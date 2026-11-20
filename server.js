require("dotenv").config();

const express = require("express");
const path = require("path");
const fs = require("fs");
const { createServer } = require("http");
const { WebSocketServer, WebSocket } = require("ws");

const {
    CommunicationIdentityClient
} = require("@azure/communication-identity");


const app = express();

const PORT = 3000;
const speakerOutputDir = path.join(__dirname, "speaker-events");
const speakerSessions = new Map();
const callStatusOutputDir = path.join(__dirname, "call-status");
const audioOutputDir = path.join(__dirname, "audio-recordings");
const callStatusSessions = new Map();
const allowedCallStatuses = new Set([
    "joining",
    "inLobby",
    "connected",
    "disconnected",
    "rejected",
    "failed"
]);

app.use(express.json({ limit: "256kb" }));


const connectionString =
    process.env.ACS_CONNECTION_STRING;


if (!connectionString) {

    console.error(
        "ERROR: ACS_CONNECTION_STRING is not configured."
    );

    process.exit(1);
}


const identityClient =
    new CommunicationIdentityClient(
        connectionString
    );


app.use(
    express.static(
        path.join(__dirname, "public")
    )
);


app.get(
    "/api/token",
    async (req, res) => {

        try {

            console.log(
                "Creating ACS identity..."
            );

            const user =
                await identityClient.createUser();

            console.log(
                "ACS user:",
                user.communicationUserId
            );

            const tokenResponse =
                await identityClient.getToken(
                    user,
                    ["voip"]
                );

            res.json({

                userId:
                    user.communicationUserId,

                token:
                    tokenResponse.token,

                expiresOn:
                    tokenResponse.expiresOn
            });

        } catch (error) {

            console.error(
                "ACS token error:",
                error
            );

            res.status(500).json({

                error:
                    error.message
            });
        }
    }
);


// --------------------------------------------------
// CALL STATUS API FOR FRONTEND INTEGRATION
// --------------------------------------------------

function saveCallStatus(session) {
    fs.mkdirSync(callStatusOutputDir, { recursive: true });
    const filePath = path.join(callStatusOutputDir, `${session.sessionId}.json`);
    fs.writeFileSync(filePath, JSON.stringify(session, null, 2), "utf8");
    return filePath;
}

// The ACS browser client posts lifecycle updates here.
app.post("/api/call-status", (req, res) => {
    try {
        const { sessionId, status, timestamp, details = {} } = req.body || {};

        if (
            typeof sessionId !== "string" ||
            !/^[a-zA-Z0-9_-]{1,100}$/.test(sessionId) ||
            !allowedCallStatuses.has(status) ||
            typeof timestamp !== "string" ||
            Number.isNaN(Date.parse(timestamp))
        ) {
            return res.status(400).json({
                error: "Invalid call status payload.",
                allowedStatuses: [...allowedCallStatuses]
            });
        }

        let session = callStatusSessions.get(sessionId);
        if (!session) {
            session = {
                sessionId,
                status,
                timestamp,
                details,
                history: []
            };
            callStatusSessions.set(sessionId, session);
        }

        const statusEvent = { status, timestamp, details };
        session.status = status;
        session.timestamp = timestamp;
        session.details = details;
        session.history.push(statusEvent);
        session.updatedAt = new Date().toISOString();

        const filePath = saveCallStatus(session);
        console.log(`[CALL STATUS] ${status} | session=${sessionId} | file=${filePath}`);

        return res.json({
            ok: true,
            sessionId,
            status: session.status,
            timestamp: session.timestamp
        });
    } catch (error) {
        console.error("Call status API error:", error);
        return res.status(500).json({ error: "Could not save call status." });
    }
});

// Frontend can poll this endpoint when it only manages one active bot session.
app.get("/api/call-status/latest", (req, res) => {
    const sessions = [...callStatusSessions.values()];
    if (!sessions.length) {
        return res.status(404).json({ status: "not_started", message: "No call status has been published yet." });
    }

    sessions.sort((a, b) => Date.parse(b.updatedAt || b.timestamp) - Date.parse(a.updatedAt || a.timestamp));
    return res.json(sessions[0]);
});

// Frontend can poll a specific meeting session using its session ID.
app.get("/api/call-status/:sessionId", (req, res) => {
    const session = callStatusSessions.get(req.params.sessionId);
    if (!session) {
        return res.status(404).json({ status: "not_found", sessionId: req.params.sessionId });
    }
    return res.json(session);
});


// --------------------------------------------------
// SPEAKER ACTIVITY EVENT STORAGE
// --------------------------------------------------

function getOrCreateSpeakerSession(sessionId) {
    if (!speakerSessions.has(sessionId)) {
        speakerSessions.set(sessionId, {
            sessionId,
            startedAt: null,
            endedAt: null,
            events: [],
            speakersById: {}
        });
    }

    return speakerSessions.get(sessionId);
}

function buildSpeakerReport(session) {
    const speakers = Object.values(session.speakersById).map(speaker => {
        const completedDuration = speaker.segments.reduce(
            (total, segment) => total + (segment.durationMs || 0),
            0
        );

        return {
            participantId: speaker.participantId,
            name: speaker.name,
            speakingTurns: speaker.segments.length,
            totalSpeakingDurationMs: completedDuration,
            currentlySpeakingSince: speaker.openStartTime,
            segments: speaker.segments
        };
    });

    return {
        sessionId: session.sessionId,
        startedAt: session.startedAt,
        endedAt: session.endedAt,
        totalSpeakers: speakers.filter(speaker =>
            speaker.speakingTurns > 0 || speaker.currentlySpeakingSince
        ).length,
        speakers,
        events: session.events,
        note: "ACS isSpeakingChanged is a dominant-speaker/speaking-state indicator. These are SDK event-observation timestamps, not word-level or exact acoustic boundaries."
    };
}

app.post("/api/speaker-events", (req, res) => {
    try {
        const { sessionId, event } = req.body || {};

        if (
            typeof sessionId !== "string" ||
            !/^[a-zA-Z0-9_-]{1,100}$/.test(sessionId) ||
            !event ||
            typeof event.type !== "string" ||
            typeof event.timestamp !== "string"
        ) {
            return res.status(400).json({
                error: "Invalid sessionId or speaker event payload."
            });
        }

        const session = getOrCreateSpeakerSession(sessionId);
        session.events.push(event);

        if (event.type === "SESSION_STARTED") {
            session.startedAt = event.timestamp;
        } else if (event.type === "SESSION_ENDED") {
            session.endedAt = event.timestamp;
        } else if (
            event.type === "SPEAKER_STARTED" ||
            event.type === "SPEAKER_STOPPED" ||
            event.type === "PARTICIPANT_JOINED" ||
            event.type === "PARTICIPANT_LEFT"
        ) {
            const participantId = event.participantId || event.participantName || "unknown";

            if (!session.speakersById[participantId]) {
                session.speakersById[participantId] = {
                    participantId,
                    name: event.participantName || "Unknown participant",
                    openStartTime: null,
                    segments: []
                };
            }

            const speaker = session.speakersById[participantId];
            speaker.name = event.participantName || speaker.name;

            if (event.type === "SPEAKER_STARTED" && !speaker.openStartTime) {
                speaker.openStartTime = event.timestamp;
            }

            if (event.type === "SPEAKER_STOPPED" && speaker.openStartTime) {
                const startMs = Date.parse(speaker.openStartTime);
                const endMs = Date.parse(event.timestamp);

                speaker.segments.push({
                    startTime: speaker.openStartTime,
                    endTime: event.timestamp,
                    durationMs: Number.isFinite(startMs) && Number.isFinite(endMs)
                        ? Math.max(0, endMs - startMs)
                        : null,
                    reason: event.reason || "speaking_state_changed"
                });

                speaker.openStartTime = null;
            }
        }

        fs.mkdirSync(speakerOutputDir, { recursive: true });
        const filePath = path.join(speakerOutputDir, `${sessionId}.json`);
        fs.writeFileSync(
            filePath,
            JSON.stringify(buildSpeakerReport(session), null, 2),
            "utf8"
        );

        console.log(
            `[SPEAKER TRACKING] ${event.type} | session=${sessionId} | file=${filePath}`
        );

        return res.json({
            ok: true,
            file: path.relative(__dirname, filePath),
            totalSpeakers: buildSpeakerReport(session).totalSpeakers
        });
    } catch (error) {
        console.error("Speaker event storage error:", error);
        return res.status(500).json({ error: "Could not save speaker event." });
    }
});


// --------------------------------------------------
// LIVE AUDIO WEBSOCKET + FINAL WAV ASSEMBLY
// --------------------------------------------------

function sendAudioSocketMessage(socket, payload) {
    if (socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify(payload));
    }
}

function createPcmWavHeader(dataBytes, sampleRate = 16000, channels = 1, bitDepth = 16) {
    const header = Buffer.alloc(44);
    const bytesPerSample = bitDepth / 8;
    const blockAlign = channels * bytesPerSample;
    const byteRate = sampleRate * blockAlign;

    header.write("RIFF", 0);
    header.writeUInt32LE(36 + dataBytes, 4);
    header.write("WAVE", 8);
    header.write("fmt ", 12);
    header.writeUInt32LE(16, 16); // PCM format chunk length
    header.writeUInt16LE(1, 20);  // WAVE_FORMAT_PCM
    header.writeUInt16LE(channels, 22);
    header.writeUInt32LE(sampleRate, 24);
    header.writeUInt32LE(byteRate, 28);
    header.writeUInt16LE(blockAlign, 32);
    header.writeUInt16LE(bitDepth, 34);
    header.write("data", 36);
    header.writeUInt32LE(dataBytes, 40);
    return header;
}

function waitForStreamOpen(stream) {
    return new Promise((resolve, reject) => {
        stream.once("open", resolve);
        stream.once("error", reject);
    });
}

function closePcmWriteStream(stream) {
    return new Promise((resolve, reject) => {
        if (stream.closed) return resolve();
        stream.once("error", reject);
        stream.end(resolve);
    });
}

function validIsoTimestamp(value) {
    return typeof value === "string" && Number.isFinite(Date.parse(value))
        ? value
        : null;
}

function buildAudioTimeline(session, finalSampleCount) {
    const connectedMs = session.meetingConnectedAtUtc
        ? Date.parse(session.meetingConnectedAtUtc)
        : NaN;
    const endedMs = session.meetingEndedAtUtc
        ? Date.parse(session.meetingEndedAtUtc)
        : NaN;
    const connectedDurationMs = Number.isFinite(connectedMs) && Number.isFinite(endedMs)
        ? Math.max(0, endedMs - connectedMs)
        : null;
    const audioDurationMs = finalSampleCount / session.sampleRate * 1000;

    return {
        inLobbyAtUtc: session.inLobbyAtUtc || null,
        meetingConnectedAtUtc: session.meetingConnectedAtUtc || null,
        audioTimelineStartAtUtc: session.meetingConnectedAtUtc || null,
        wavTimelineStartAtUtc: session.meetingConnectedAtUtc || null,
        captureRequestedAtUtc: session.captureRequestedAtUtc || null,
        audioCapturePipelineReadyAtUtc: session.audioCapturePipelineReadyAtUtc || null,
        audioProcessorStartedAtUtc: session.audioProcessorStartedAtUtc || null,
        meetingEndedAtUtc: session.meetingEndedAtUtc || null,
        connectedDurationMs,
        connectedDurationSeconds: connectedDurationMs === null
            ? null
            : Number((connectedDurationMs / 1000).toFixed(3)),
        audioDurationMs: Number(audioDurationMs.toFixed(3)),
        audioDurationSeconds: Number((audioDurationMs / 1000).toFixed(3)),
        audioVsConnectedDurationDeltaMs: connectedDurationMs === null
            ? null
            : Number((audioDurationMs - connectedDurationMs).toFixed(3)),
        targetSampleCount: session.targetSampleCount,
        utcElapsedMs: session.utcElapsedMs,
        monotonicElapsedMs: session.monotonicElapsedMs,
        wallClockVsMonotonicDeltaMs: session.wallClockVsMonotonicDeltaMs,
        workletReportedSampleCount: session.workletReportedSampleCount,
        audioLastSampleIndexSent: session.audioLastSampleIndexSent,
        audioSamplesSent: session.audioSamplesSent,
        pendingSuspension: session.pendingSuspension,
        finalSampleCount,
        sampleCountBeforeReconciliation: session.sampleCountBeforeReconciliation,
        rawSamplesReceived: session.rawSamplesReceived,
        serverInsertedGapSamples: session.serverInsertedGapSamples,
        samplesPaddedAtEnd: session.samplesPaddedAtEnd,
        samplesTrimmedAtEnd: session.samplesTrimmedAtEnd,
        sampleCountDeltaBeforeReconciliation: session.sampleCountBeforeReconciliation === null
            ? null
            : session.targetSampleCount === null
                ? null
                : session.targetSampleCount - session.sampleCountBeforeReconciliation,
        sampleIndexProtocolVersion: 2,
        sampleIndexGaps: session.sampleIndexGaps,
        declaredAudioGaps: session.declaredAudioGaps,
        declaredAudioGapSamples: session.declaredAudioGaps.reduce((sum, gap) => sum + (Number(gap.sampleCount) || 0), 0),
        sequenceDiscontinuities: session.sequenceDiscontinuities,
        audioContextEvents: session.audioContextEvents,
        workletStats: session.workletStats,
        chunkManifestCount: session.chunkManifest.length,
        initialSilenceSamples: session.audioInitialSilenceSamples || 0,
        initialSilenceMs: Number((((session.audioInitialSilenceSamples || 0) / session.sampleRate) * 1000).toFixed(3))
    };
}

async function writeWavFromPcm(pcmPath, wavPath, dataBytes, sampleRate, channels, bitDepth) {
    if (dataBytes > 0xFFFFFFFF - 36) {
        throw new Error("PCM data exceeds classic WAV 4 GiB size limit; use RF64 for longer recordings.");
    }

    await new Promise((resolve, reject) => {
        const output = fs.createWriteStream(wavPath);
        const input = fs.createReadStream(pcmPath);
        let settled = false;
        const fail = (error) => {
            if (settled) return;
            settled = true;
            output.destroy();
            input.destroy();
            reject(error);
        };

        output.once("error", fail);
        input.once("error", fail);
        output.once("finish", () => {
            if (settled) return;
            settled = true;
            resolve();
        });

        output.write(createPcmWavHeader(dataBytes, sampleRate, channels, bitDepth));
        input.pipe(output);
    });
}

async function appendZeroSamplesToFile(filePath, sampleCount) {
    if (!Number.isSafeInteger(sampleCount) || sampleCount <= 0) return;

    const handle = await fs.promises.open(filePath, "a");
    const zeroBlock = Buffer.alloc(64 * 1024);
    let remainingBytes = sampleCount * 2;
    try {
        while (remainingBytes > 0) {
            const length = Math.min(remainingBytes, zeroBlock.length);
            await handle.write(zeroBlock, 0, length);
            remainingBytes -= length;
        }
    } finally {
        await handle.close();
    }
}

async function reconcilePcmToTargetSamples(session) {
    const before = await fs.promises.stat(session.pcmPath);
    const bytesPerSample = session.bitDepth / 8;
    const actualSamples = Math.floor(before.size / bytesPerSample);
    session.sampleCountBeforeReconciliation = actualSamples;

    if (session.targetSampleCount === null) return actualSamples;

    if (!Number.isSafeInteger(session.targetSampleCount) || session.targetSampleCount < 0 || session.targetSampleCount > 2_000_000_000) {
        throw new Error("Invalid targetSampleCount; expected an integer between 0 and 2,000,000,000.");
    }

    if (actualSamples < session.targetSampleCount) {
        const missing = session.targetSampleCount - actualSamples;
        await appendZeroSamplesToFile(session.pcmPath, missing);
        session.samplesPaddedAtEnd = missing;
        session.declaredAudioGaps.push({
            reason: "end_of_call_sample_reconciliation",
            startSample: actualSamples,
            sampleCount: missing,
            durationMs: Number((missing / session.sampleRate * 1000).toFixed(3)),
            startUtc: null,
            endUtc: session.meetingEndedAtUtc
        });
    } else if (actualSamples > session.targetSampleCount) {
        const trimmed = actualSamples - session.targetSampleCount;
        await fs.promises.truncate(session.pcmPath, session.targetSampleCount * bytesPerSample);
        session.samplesTrimmedAtEnd = trimmed;
    }

    return session.targetSampleCount;
}

async function finalizeAudioSession(session, reason = "meeting_ended") {
    if (!session || session.finalized) return session?.finalResult || null;
    await closePcmWriteStream(session.pcmStream);
    const finalSampleCount = await reconcilePcmToTargetSamples(session);
    const stats = await fs.promises.stat(session.pcmPath);
    const audioBytes = stats.size;
    let wavPath = null;

    if (audioBytes > 0) {
        await writeWavFromPcm(
            session.pcmPath,
            session.wavPath,
            audioBytes,
            session.sampleRate,
            session.channels,
            session.bitDepth
        );
        wavPath = path.relative(__dirname, session.wavPath).split(path.sep).join("/");
    }

    const durationSeconds = finalSampleCount / session.sampleRate;
    const result = {
        sessionId: session.sessionId,
        wavPath,
        pcmPath: path.relative(__dirname, session.pcmPath).split(path.sep).join("/"),
        chunks: session.chunks,
        audioBytes,
        durationSeconds: Number(durationSeconds.toFixed(3)),
        sampleRate: session.sampleRate,
        channels: session.channels,
        bitDepth: session.bitDepth,
        reason,
        timeline: buildAudioTimeline(session, finalSampleCount),
        chunkManifest: session.chunkManifest,
        finalizedAt: new Date().toISOString()
    };

    const metadataPath = path.join(audioOutputDir, `${session.sessionId}.json`);
    await fs.promises.writeFile(metadataPath, JSON.stringify(result, null, 2), "utf8");
    session.finalResult = result;
    session.finalized = true;

    console.log(
        `[LIVE AUDIO] Finalized session=${session.sessionId} chunks=${session.chunks} receivedSamples=${session.rawSamplesReceived} finalSamples=${finalSampleCount} duration=${result.durationSeconds}s padded=${session.samplesPaddedAtEnd} trimmed=${session.samplesTrimmedAtEnd} wav=${wavPath || "(empty; no WAV created)"}`
    );
    return result;
}

const httpServer = createServer(app);
const audioWebSocketServer = new WebSocketServer({
    server: httpServer,
    path: "/ws/audio",
    maxPayload: 128 * 1024
});

const PCM_FRAME_MAGIC = 0x314d4350; // "PCM1".
const PCM_FRAME_HEADER_BYTES = 20;

function appendAudioBuffer(stream, buffer) {
    return new Promise((resolve, reject) => {
        stream.write(buffer, (error) => error ? reject(error) : resolve());
    });
}

async function writeSilenceSamples(session, sampleCount, reason, detectedAtUtc = new Date().toISOString()) {
    if (!Number.isSafeInteger(sampleCount) || sampleCount <= 0) return;

    const startSample = session.samplesWritten;
    const zeroBlock = Buffer.alloc(64 * 1024);
    let remainingBytes = sampleCount * 2;
    while (remainingBytes > 0) {
        const length = Math.min(remainingBytes, zeroBlock.length);
        await appendAudioBuffer(session.pcmStream, zeroBlock.subarray(0, length));
        remainingBytes -= length;
    }

    session.samplesWritten += sampleCount;
    session.serverInsertedGapSamples += sampleCount;
    session.sampleIndexGaps.push({
        reason,
        startSample,
        sampleCount,
        durationMs: Number((sampleCount / session.sampleRate * 1000).toFixed(3)),
        detectedAtUtc
    });
}

audioWebSocketServer.on("connection", (socket) => {
    let audioSession = null;
    let messageQueue = Promise.resolve();

    console.log("[LIVE AUDIO] Browser WebSocket connected (sample-index protocol v2).");

    socket.on("message", (data, isBinary) => {
        messageQueue = messageQueue.then(async () => {
            if (isBinary) {
                if (!audioSession) {
                    throw new Error("Received PCM bytes before an audio session start message.");
                }

                const frame = Buffer.isBuffer(data) ? data : Buffer.from(data);
                if (frame.length < PCM_FRAME_HEADER_BYTES + 2) {
                    throw new Error("Invalid PCM frame: header or PCM payload is missing.");
                }
                if (frame.readUInt32LE(0) !== PCM_FRAME_MAGIC) {
                    throw new Error("Invalid PCM frame magic. Client/server protocol versions may not match.");
                }

                const sequence = frame.readUInt32LE(4);
                const sampleOffsetBigInt = frame.readBigUInt64LE(8);
                const sampleCount = frame.readUInt32LE(16);
                const sampleOffset = Number(sampleOffsetBigInt);
                const pcm = frame.subarray(PCM_FRAME_HEADER_BYTES);

                if (!Number.isSafeInteger(sampleOffset) || sampleCount <= 0 || pcm.length !== sampleCount * 2) {
                    throw new Error("Invalid indexed PCM frame: sample offset/count does not match PCM16 payload size.");
                }

                if (sequence < audioSession.nextSequenceExpected) {
                    throw new Error(`Out-of-order/duplicate PCM sequence ${sequence}; expected ${audioSession.nextSequenceExpected}.`);
                }
                if (sequence > audioSession.nextSequenceExpected) {
                    audioSession.sequenceDiscontinuities.push({
                        expectedSequence: audioSession.nextSequenceExpected,
                        receivedSequence: sequence,
                        missingSequenceCount: sequence - audioSession.nextSequenceExpected,
                        detectedAtUtc: new Date().toISOString()
                    });
                }
                audioSession.nextSequenceExpected = sequence + 1;

                if (sampleOffset > audioSession.samplesWritten) {
                    const gapSamples = sampleOffset - audioSession.samplesWritten;
                    await writeSilenceSamples(audioSession, gapSamples, "missing_sample_index_range_before_chunk");
                } else if (sampleOffset < audioSession.samplesWritten) {
                    throw new Error(`Overlapping PCM frame: firstSample=${sampleOffset}, nextExpectedSample=${audioSession.samplesWritten}.`);
                }

                await appendAudioBuffer(audioSession.pcmStream, pcm);
                audioSession.samplesWritten += sampleCount;
                audioSession.rawSamplesReceived += sampleCount;
                audioSession.chunks += 1;
                if (audioSession.chunkManifest.length < 100000) {
                    audioSession.chunkManifest.push({
                        sequence,
                        firstSampleIndex: sampleOffset,
                        sampleCount,
                        endSampleIndex: sampleOffset + sampleCount
                    });
                }

                if (audioSession.chunks === 1 || audioSession.chunks % 10 === 0) {
                    console.log(`[LIVE AUDIO] Received chunk=${audioSession.chunks} seq=${sequence} firstSample=${sampleOffset} samples=${sampleCount} endSample=${audioSession.samplesWritten} session=${audioSession.sessionId}`);
                }
                return;
            }

            let message;
            try {
                message = JSON.parse(data.toString());
            } catch {
                throw new Error("Invalid audio WebSocket JSON control message.");
            }

            if (message.type === "start") {
                if (audioSession) throw new Error("An audio session is already active on this WebSocket.");

                const {
                    sessionId,
                    sampleRate,
                    channels,
                    bitDepth,
                    protocolVersion,
                    inLobbyAtUtc,
                    meetingConnectedAtUtc,
                    captureRequestedAtUtc,
                    audioTimelineStartAtUtc
                } = message;

                if (
                    typeof sessionId !== "string" ||
                    !/^[a-zA-Z0-9_-]{1,100}$/.test(sessionId) ||
                    sampleRate !== 16000 ||
                    channels !== 1 ||
                    bitDepth !== 16 ||
                    protocolVersion !== 2
                ) {
                    throw new Error("Invalid audio start message. Expected protocol v2, indexed PCM16 mono 16 kHz.");
                }

                fs.mkdirSync(audioOutputDir, { recursive: true });
                const pcmPath = path.join(audioOutputDir, `${sessionId}.pcm`);
                const wavPath = path.join(audioOutputDir, `${sessionId}.wav`);
                const pcmStream = fs.createWriteStream(pcmPath, { flags: "wx" });
                await waitForStreamOpen(pcmStream);

                audioSession = {
                    sessionId,
                    protocolVersion,
                    sampleRate,
                    channels,
                    bitDepth,
                    pcmPath,
                    wavPath,
                    pcmStream,
                    chunks: 0,
                    samplesWritten: 0,
                    rawSamplesReceived: 0,
                    serverInsertedGapSamples: 0,
                    nextSequenceExpected: 0,
                    sampleIndexGaps: [],
                    declaredAudioGaps: [],
                    sequenceDiscontinuities: [],
                    chunkManifest: [],
                    audioContextEvents: [],
                    workletStats: [],
                    targetSampleCount: null,
                    utcElapsedMs: null,
                    monotonicElapsedMs: null,
                    wallClockVsMonotonicDeltaMs: null,
                    workletReportedSampleCount: null,
                    audioLastSampleIndexSent: null,
                    audioSamplesSent: null,
                    pendingSuspension: null,
                    sampleCountBeforeReconciliation: null,
                    samplesPaddedAtEnd: 0,
                    samplesTrimmedAtEnd: 0,
                    inLobbyAtUtc: validIsoTimestamp(inLobbyAtUtc),
                    meetingConnectedAtUtc: validIsoTimestamp(meetingConnectedAtUtc),
                    audioTimelineStartAtUtc: validIsoTimestamp(audioTimelineStartAtUtc) || validIsoTimestamp(meetingConnectedAtUtc),
                    captureRequestedAtUtc: validIsoTimestamp(captureRequestedAtUtc),
                    audioCapturePipelineReadyAtUtc: null,
                    audioProcessorStartedAtUtc: null,
                    meetingEndedAtUtc: null,
                    audioInitialSilenceSamples: 0,
                    finalized: false,
                    finalResult: null
                };

                console.log(`[LIVE AUDIO] Started session=${sessionId} protocol=2 format=PCM16/${sampleRate}/mono`);
                sendAudioSocketMessage(socket, {
                    type: "started",
                    protocolVersion: 2,
                    sessionId,
                    sampleRate,
                    channels,
                    bitDepth
                });
                return;
            }

            if (!audioSession) {
                throw new Error(`Received ${message.type} before audio session start.`);
            }

            if (message.type === "timeline") {
                audioSession.audioCapturePipelineReadyAtUtc = validIsoTimestamp(message.audioCapturePipelineReadyAtUtc)
                    || audioSession.audioCapturePipelineReadyAtUtc;
                audioSession.audioProcessorStartedAtUtc = validIsoTimestamp(message.audioProcessorStartedAtUtc)
                    || audioSession.audioProcessorStartedAtUtc;
                if (Number.isSafeInteger(message.audioInitialSilenceSamples) && message.audioInitialSilenceSamples >= 0) {
                    audioSession.audioInitialSilenceSamples = message.audioInitialSilenceSamples;
                }
                sendAudioSocketMessage(socket, { type: "timeline_saved", sessionId: audioSession.sessionId });
                return;
            }

            if (message.type === "gap") {
                const startSample = Number(message.startSample);
                const sampleCount = Number(message.sampleCount);
                if (!Number.isSafeInteger(startSample) || startSample < 0 || !Number.isSafeInteger(sampleCount) || sampleCount <= 0) {
                    throw new Error("Invalid declared audio gap metadata.");
                }
                audioSession.declaredAudioGaps.push({
                    reason: typeof message.reason === "string" ? message.reason.slice(0, 120) : "audio_gap",
                    startSample,
                    sampleCount,
                    startUtc: validIsoTimestamp(message.startUtc),
                    endUtc: validIsoTimestamp(message.endUtc),
                    durationMs: Number.isFinite(Number(message.durationMs)) ? Number(Number(message.durationMs).toFixed(3)) : null
                });
                return;
            }

            if (message.type === "contextState") {
                if (audioSession.audioContextEvents.length < 2000) {
                    audioSession.audioContextEvents.push({
                        state: String(message.state || "unknown").slice(0, 40),
                        previousState: message.previousState ? String(message.previousState).slice(0, 40) : null,
                        timestampUtc: validIsoTimestamp(message.timestampUtc),
                        performanceTimeMs: Number.isFinite(Number(message.performanceTimeMs)) ? Number(message.performanceTimeMs) : null,
                        audioContextTimeSeconds: Number.isFinite(Number(message.audioContextTimeSeconds)) ? Number(message.audioContextTimeSeconds) : null
                    });
                }
                return;
            }

            if (message.type === "stats") {
                if (audioSession.workletStats.length < 2000 && message.stats && typeof message.stats === "object") {
                    audioSession.workletStats.push({
                        ...message.stats,
                        timestampUtc: validIsoTimestamp(message.timestampUtc) || new Date().toISOString()
                    });
                }
                return;
            }

            if (message.type === "end") {
                audioSession.meetingEndedAtUtc = validIsoTimestamp(message.meetingEndedAtUtc) || new Date().toISOString();
                audioSession.inLobbyAtUtc = validIsoTimestamp(message.inLobbyAtUtc) || audioSession.inLobbyAtUtc;
                audioSession.meetingConnectedAtUtc = validIsoTimestamp(message.meetingConnectedAtUtc) || audioSession.meetingConnectedAtUtc;
                audioSession.audioTimelineStartAtUtc = audioSession.meetingConnectedAtUtc || audioSession.audioTimelineStartAtUtc;
                audioSession.captureRequestedAtUtc = validIsoTimestamp(message.captureRequestedAtUtc) || audioSession.captureRequestedAtUtc;
                audioSession.audioCapturePipelineReadyAtUtc = validIsoTimestamp(message.audioCapturePipelineReadyAtUtc) || audioSession.audioCapturePipelineReadyAtUtc;
                audioSession.audioProcessorStartedAtUtc = validIsoTimestamp(message.audioProcessorStartedAtUtc) || audioSession.audioProcessorStartedAtUtc;
                if (Number.isSafeInteger(message.audioInitialSilenceSamples) && message.audioInitialSilenceSamples >= 0) {
                    audioSession.audioInitialSilenceSamples = message.audioInitialSilenceSamples;
                }

                if (Number.isSafeInteger(message.targetSampleCount) && message.targetSampleCount >= 0) {
                    if (message.targetSampleCount > 2_000_000_000) {
                        throw new Error("targetSampleCount exceeds supported classic WAV duration limit.");
                    }
                    audioSession.targetSampleCount = message.targetSampleCount;
                }
                audioSession.utcElapsedMs = Number.isFinite(Number(message.utcElapsedMs)) ? Number(message.utcElapsedMs) : null;
                audioSession.monotonicElapsedMs = Number.isFinite(Number(message.monotonicElapsedMs)) ? Number(message.monotonicElapsedMs) : null;
                audioSession.wallClockVsMonotonicDeltaMs = Number.isFinite(Number(message.wallClockVsMonotonicDeltaMs)) ? Number(message.wallClockVsMonotonicDeltaMs) : null;
                audioSession.workletReportedSampleCount = Number.isSafeInteger(message.workletReportedSampleCount) ? message.workletReportedSampleCount : null;
                audioSession.audioLastSampleIndexSent = Number.isSafeInteger(message.audioLastSampleIndexSent) ? message.audioLastSampleIndexSent : null;
                audioSession.audioSamplesSent = Number.isSafeInteger(message.audioSamplesSent) ? message.audioSamplesSent : null;
                audioSession.pendingSuspension = message.pendingSuspension && typeof message.pendingSuspension === "object"
                    ? message.pendingSuspension
                    : null;
                if (Array.isArray(message.contextStateEvents)) audioSession.audioContextEvents = message.contextStateEvents.slice(0, 2000);
                if (Array.isArray(message.timelineGaps)) audioSession.declaredAudioGaps = message.timelineGaps.slice(0, 2000);
                if (Array.isArray(message.workletStats)) audioSession.workletStats = message.workletStats.slice(0, 2000);

                const result = await finalizeAudioSession(audioSession, message.reason || "meeting_ended");
                audioSession = null;
                sendAudioSocketMessage(socket, { type: "finalized", ...result });
                return;
            }

            throw new Error(`Unknown audio WebSocket message type: ${String(message.type)}`);
        }).catch((error) => {
            console.error("[LIVE AUDIO] WebSocket processing error:", error);
            sendAudioSocketMessage(socket, { type: "error", message: error.message || "Audio WebSocket processing failed." });
        });
    });

    socket.on("close", () => {
        messageQueue = messageQueue.then(async () => {
            if (!audioSession) return;
            try {
                audioSession.meetingEndedAtUtc = audioSession.meetingEndedAtUtc || new Date().toISOString();
                await finalizeAudioSession(audioSession, "websocket_closed_unexpectedly");
            } catch (error) {
                console.error("[LIVE AUDIO] Could not salvage interrupted audio session:", error);
            }
            audioSession = null;
        });
    });

    socket.on("error", (error) => {
        console.error("[LIVE AUDIO] WebSocket transport error:", error);
    });
});

httpServer.listen(
    PORT,
    () => {
        console.log(`ACS backend running at http://localhost:${PORT}`);
        console.log(`Audio WebSocket listening at ws://localhost:${PORT}/ws/audio`);
    }
);