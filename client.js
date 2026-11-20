import {
    CallClient
} from "@azure/communication-calling";

import {
    AzureCommunicationTokenCredential
} from "@azure/communication-common";


let callClient = null;
let callAgent = null;
let call = null;
let tokenCredential = null;


// --------------------------------------------------
// LIVE REMOTE AUDIO CAPTURE (PCM16 / 16 kHz / MONO)
// --------------------------------------------------

const AUDIO_SAMPLE_RATE = 16000;
const AUDIO_FRAME_MAGIC = 0x314d4350; // "PCM1" in little-endian bytes.
const AUDIO_FRAME_HEADER_BYTES = 20;
let audioContext = null;
let audioWorkletNode = null;
let audioSilentGain = null;
let audioSocket = null;
let audioStartPromise = null;
let audioStopPromise = null;
let audioFlushResolver = null;
let audioChunkCount = 0;
let audioBytesSent = 0;
let audioSamplesSent = 0;
let audioLastSampleIndexSent = 0;
let audioWorkletTotalSampleCount = null;
let audioContextStateHandler = null;
let audioContextLastState = null;
let audioContextSuspendedAtPerfMs = null;
let audioContextSuspendedAtUtc = null;
let audioContextEvents = [];
let audioTimelineGaps = [];
let audioWorkletStats = [];
const audioSourcesByRemoteStream = new Map();
const audioStreamAttachPromises = new Map();
let remoteAudioStreamsUpdatedHandler = null;


// --------------------------------------------------
// SPEAKER ACTIVITY TRACKING (ACS SDK)
// --------------------------------------------------

let speakerSessionId = null;
let speakerSessionEnded = false;
let speakerEventQueue = Promise.resolve();
const participantSpeakerMonitors = new Map();

// Meeting lifecycle status for the frontend integration.
let callStatusQueue = Promise.resolve();
let lastObservedCallState = null;
let hasEnteredLobby = false;
let hasConnectedToMeeting = false;
let terminalCallStatusPublished = false;

// UTC/monotonic anchors for matching the WAV timeline to the ACS call.
let inLobbyAtUtc = null;
let callConnectedAtUtc = null;
let callConnectedPerfMs = null;
let meetingEndedAtUtc = null;
let meetingEndedPerfMs = null;
let captureRequestedAtUtc = null;
let audioCapturePipelineReadyAtUtc = null;
let audioProcessorStartedAtUtc = null;
let audioInitialSilenceSamples = 0;

// Auto-leave is enabled by default and can be changed in the page checkbox.
const AUTO_LEAVE_STORAGE_KEY = "acs.autoLeaveWhenEveryoneLeaves";
let autoLeaveWhenEmpty = true;
let hasSeenRemoteParticipant = false;
let autoLeaveTriggered = false;
let autoLeaveCheckTimer = null;
let leaveInProgress = false;

function publishCallStatus(status, details = {}) {
    if (!speakerSessionId) {
        console.warn("Cannot publish call status: no session ID.", status);
        return callStatusQueue;
    }

    const payload = {
        sessionId: speakerSessionId,
        status,
        timestamp: new Date().toISOString(),
        details
    };

    // Keep status events ordered, just like speaker activity events.
    callStatusQueue = callStatusQueue
        .then(async () => {
            const response = await fetch("/api/call-status", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const body = await response.text();
                throw new Error(`Call status API failed (${response.status}): ${body}`);
            }

            const result = await response.json();
            console.log("[CALL STATUS API]", result.status, result.timestamp);
        })
        .catch(error => {
            console.error("Could not publish call status:", error);
        });

    return callStatusQueue;
}

function getCallEndReasonDetails(activeCall) {
    const reason = activeCall?.callEndReason;
    if (!reason) return null;

    return {
        code: reason.code ?? null,
        subCode: reason.subCode ?? null,
        message: reason.message ?? null,
        resultCategories: reason.resultCategories ?? null
    };
}

async function handleCallStateChange() {
    if (!call) return;

    const state = call.state;
    const previousCallState = lastObservedCallState;
    lastObservedCallState = state;

    console.log("ACS call state:", state);

    if (state === "InLobby") {
        hasEnteredLobby = true;
        if (!inLobbyAtUtc) inLobbyAtUtc = new Date().toISOString();

        await publishCallStatus("inLobby", {
            callState: state,
            inLobbyAtUtc,
            message: "Bot is waiting in the Teams meeting lobby."
        });
        setStatus("Waiting in Teams lobby...");
        return;
    }

    if (state === "Connected") {
        hasConnectedToMeeting = true;
        terminalCallStatusPublished = false;

        // Record the exact first Connected transition (not the lobby arrival).
        // Start the PCM timeline immediately, even if ACS has not exposed a
        // remoteAudioStream yet. The worklet fills that initial period with
        // silence, then accepts audio once the remote stream is attached.
        if (!callConnectedAtUtc) {
            callConnectedAtUtc = new Date().toISOString();
            callConnectedPerfMs = performance.now();
            audioCapturePipelineReadyAtUtc = null;
            audioProcessorStartedAtUtc = null;
            audioInitialSilenceSamples = 0;
            hasSeenRemoteParticipant = (call.remoteParticipants || []).length > 0;

            queueSpeakerEvent({
                type: "CALL_CONNECTED",
                timestamp: callConnectedAtUtc,
                inLobbyAtUtc,
                callConnectedAtUtc,
                joinToConnectedDeltaMs: inLobbyAtUtc
                    ? Math.max(0, Date.parse(callConnectedAtUtc) - Date.parse(inLobbyAtUtc))
                    : null
            });

            console.log("[CALL TIMELINE] Bot admitted/connected:", {
                inLobbyAtUtc,
                callConnectedAtUtc,
                lobbyWaitMs: inLobbyAtUtc
                    ? Date.parse(callConnectedAtUtc) - Date.parse(inLobbyAtUtc)
                    : null
            });

            void startPcmCapturePipeline().catch(error => {
                console.error("[LIVE AUDIO] Could not start capture at Connected:", error);
                setStatus(`Connected, but audio capture startup failed: ${error.message || error}`);
            });
        }

        await publishCallStatus("connected", {
            callState: state,
            inLobbyAtUtc,
            callConnectedAtUtc,
            message: "Bot has been admitted and connected to the meeting."
        });
        setStatus("Connected to Teams meeting; audio timeline started.");
        scheduleAutoLeaveCheck("call_connected_roster_check");
        return;
    }

    if (state === "Disconnected") {
        if (terminalCallStatusPublished) return;
        terminalCallStatusPublished = true;
        if (!meetingEndedAtUtc) meetingEndedAtUtc = new Date().toISOString();
        if (meetingEndedPerfMs === null) meetingEndedPerfMs = performance.now();

        // Flush the last PCM samples and wait for the server to finalize the WAV.
        await stopLiveAudioCapture("call_disconnected");

        if (remoteAudioStreamsUpdatedHandler) {
            try { call.off("remoteAudioStreamsUpdated", remoteAudioStreamsUpdatedHandler); } catch { }
            remoteAudioStreamsUpdatedHandler = null;
        }

        const callEndReason = getCallEndReasonDetails(call);
        let finalStatus;
        let message;

        if (hasConnectedToMeeting) {
            finalStatus = "disconnected";
            message = "Bot disconnected from the Teams meeting.";
        } else if (hasEnteredLobby) {
            finalStatus = "rejected";
            message = "Bot was not admitted to the meeting before the call ended.";
        } else {
            finalStatus = "failed";
            message = "Call disconnected before reaching the Teams lobby or connecting.";
        }

        await publishCallStatus(finalStatus, {
            callState: state,
            previousCallState,
            inLobbyAtUtc,
            callConnectedAtUtc,
            meetingEndedAtUtc,
            callEndReason,
            classification: finalStatus === "rejected"
                ? "not_admitted_before_disconnect"
                : finalStatus,
            message
        });

        await endSpeakerTrackingSession("call_disconnected");
        setStatus(message);
        joinButton.disabled = false;
        leaveButton.disabled = true;
        return;
    }

    if (["None", "Connecting", "Ringing", "EarlyMedia", "Disconnecting"].includes(state)) {
        await publishCallStatus("joining", {
            callState: state,
            inLobbyAtUtc,
            message: "Bot is attempting to join the Teams meeting."
        });
        setStatus(`Joining Teams meeting (${state})...`);
    }
}

function readAutoLeavePreference() {
    try {
        const stored = window.localStorage.getItem(AUTO_LEAVE_STORAGE_KEY);
        return stored === null ? true : stored === "true";
    } catch (error) {
        console.warn("Could not read auto-leave setting; defaulting to enabled:", error);
        return true;
    }
}

function installAutoLeaveSetting() {
    autoLeaveWhenEmpty = readAutoLeavePreference();
    if (document.getElementById("acs-auto-leave-setting")) return;

    const label = document.createElement("label");
    label.id = "acs-auto-leave-setting";
    label.style.display = "inline-flex";
    label.style.alignItems = "center";
    label.style.gap = "8px";
    label.style.margin = "10px 0";
    label.style.fontSize = "14px";
    label.style.cursor = "pointer";
    label.style.userSelect = "none";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = autoLeaveWhenEmpty;
    checkbox.setAttribute("aria-label", "Automatically leave when everyone else leaves");
    checkbox.addEventListener("change", () => {
        autoLeaveWhenEmpty = checkbox.checked;
        try {
            window.localStorage.setItem(AUTO_LEAVE_STORAGE_KEY, String(autoLeaveWhenEmpty));
        } catch (error) {
            console.warn("Could not persist auto-leave setting:", error);
        }
        console.log("[AUTO LEAVE] Setting changed:", autoLeaveWhenEmpty);
        if (autoLeaveWhenEmpty) scheduleAutoLeaveCheck("setting_enabled");
    });

    const text = document.createElement("span");
    text.textContent = "Automatically leave when everyone else leaves";
    label.append(checkbox, text);

    const host = leaveButton?.parentElement || statusElement?.parentElement || document.body;
    host.appendChild(label);
}

function scheduleAutoLeaveCheck(reason) {
    if (!autoLeaveWhenEmpty || autoLeaveTriggered || !call) return;
    if (autoLeaveCheckTimer) clearTimeout(autoLeaveCheckTimer);

    // A short settle window lets the SDK update its remoteParticipants collection.
    autoLeaveCheckTimer = setTimeout(async () => {
        autoLeaveCheckTimer = null;
        if (!autoLeaveWhenEmpty || autoLeaveTriggered || !call || call.state !== "Connected") return;

        const remoteParticipants = Array.from(call.remoteParticipants || []);
        if (remoteParticipants.length > 0) {
            hasSeenRemoteParticipant = true;
            return;
        }

        // Do not leave just because the bot was admitted before anyone else arrived.
        if (!hasSeenRemoteParticipant) return;

        autoLeaveTriggered = true;
        const leftAtUtc = new Date().toISOString();
        meetingEndedAtUtc = meetingEndedAtUtc || leftAtUtc;
        if (meetingEndedPerfMs === null) meetingEndedPerfMs = performance.now();
        console.log("[AUTO LEAVE] All remote participants have left; bot is leaving.", {
            reason,
            meetingEndedAtUtc
        });
        setStatus("Everyone has left the meeting. Finalizing audio and leaving...");
        await leaveMeeting("all_remote_participants_left");
    }, 300);
}

function createSpeakerSessionId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
        return window.crypto.randomUUID();
    }

    return `speaker_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function getParticipantKey(participant) {
    if (participant.identifier) {
        return JSON.stringify(participant.identifier);
    }

    return participant.displayName || `participant_${participantSpeakerMonitors.size + 1}`;
}

function queueSpeakerEvent(event) {
    if (!speakerSessionId) {
        console.warn("Speaker event skipped: no active speaker session.", event);
        return speakerEventQueue;
    }

    const payload = {
        sessionId: speakerSessionId,
        event
    };

    // Serialize requests so START/STOP events reach the local server in order.
    speakerEventQueue = speakerEventQueue
        .then(async () => {
            const response = await fetch("/api/speaker-events", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const message = await response.text();
                throw new Error(`Speaker event save failed (${response.status}): ${message}`);
            }

            const result = await response.json();
            console.log("Speaker event saved:", event.type, result.file);
        })
        .catch(error => {
            console.error("Could not persist speaker event:", error);
        });

    return speakerEventQueue;
}

function startSpeakerTrackingSession() {
    speakerSessionId = createSpeakerSessionId();
    speakerSessionEnded = false;
    participantSpeakerMonitors.clear();
    callStatusQueue = Promise.resolve();
    lastObservedCallState = null;
    hasEnteredLobby = false;
    hasConnectedToMeeting = false;
    terminalCallStatusPublished = false;
    inLobbyAtUtc = null;
    callConnectedAtUtc = null;
    callConnectedPerfMs = null;
    meetingEndedAtUtc = null;
    meetingEndedPerfMs = null;
    captureRequestedAtUtc = null;
    audioCapturePipelineReadyAtUtc = null;
    audioProcessorStartedAtUtc = null;
    audioInitialSilenceSamples = 0;
    audioChunkCount = 0;
    audioBytesSent = 0;
    audioSamplesSent = 0;
    audioLastSampleIndexSent = 0;
    audioContextLastState = audioContext ? audioContext.state : null;
    audioContextSuspendedAtPerfMs = null;
    audioContextSuspendedAtUtc = null;
    audioContextEvents = [];
    audioTimelineGaps = [];
    audioWorkletStats = [];
    hasSeenRemoteParticipant = false;
    autoLeaveTriggered = false;
    leaveInProgress = false;
    if (autoLeaveCheckTimer) {
        clearTimeout(autoLeaveCheckTimer);
        autoLeaveCheckTimer = null;
    }

    console.log("Speaker tracking session:", speakerSessionId);
    publishCallStatus("joining", {
        callState: "Starting",
        message: "Bot is preparing to join the Teams meeting."
    });
    queueSpeakerEvent({
        type: "SESSION_STARTED",
        timestamp: new Date().toISOString()
    });
}

function monitorRemoteParticipant(participant) {
    const participantKey = getParticipantKey(participant);

    if (participantSpeakerMonitors.has(participantKey)) {
        return;
    }

    const monitor = {
        participant,
        isSpeaking: Boolean(participant.isSpeaking),
        listener: null
    };

    monitor.listener = () => {
        const nextIsSpeaking = Boolean(participant.isSpeaking);

        if (nextIsSpeaking === monitor.isSpeaking) {
            return;
        }

        monitor.isSpeaking = nextIsSpeaking;

        const event = {
            type: nextIsSpeaking ? "SPEAKER_STARTED" : "SPEAKER_STOPPED",
            timestamp: new Date().toISOString(),
            participantId: participantKey,
            participantName: participant.displayName || "Unknown participant"
        };

        console.log("[ACS SPEAKER ACTIVITY]", event);
        queueSpeakerEvent(event);
    };

    participantSpeakerMonitors.set(participantKey, monitor);
    participant.on("isSpeakingChanged", monitor.listener);

    queueSpeakerEvent({
        type: "PARTICIPANT_JOINED",
        timestamp: new Date().toISOString(),
        participantId: participantKey,
        participantName: participant.displayName || "Unknown participant"
    });

    console.log("Speaker activity listener attached:", {
        participantId: participantKey,
        displayName: participant.displayName,
        initialIsSpeaking: monitor.isSpeaking
    });

    // Account for a participant who is already marked as the active speaker
    // when our listener is attached.
    if (monitor.isSpeaking) {
        queueSpeakerEvent({
            type: "SPEAKER_STARTED",
            timestamp: new Date().toISOString(),
            participantId: participantKey,
            participantName: participant.displayName || "Unknown participant",
            reason: "already_speaking_when_monitor_attached"
        });
    }
}

function stopMonitoringRemoteParticipant(participant, reason = "participant_left") {
    const participantKey = getParticipantKey(participant);
    const monitor = participantSpeakerMonitors.get(participantKey);

    if (!monitor) {
        return Promise.resolve();
    }

    if (monitor.isSpeaking) {
        monitor.isSpeaking = false;
        queueSpeakerEvent({
            type: "SPEAKER_STOPPED",
            timestamp: new Date().toISOString(),
            participantId: participantKey,
            participantName: participant.displayName || "Unknown participant",
            reason
        });
    }

    participant.off("isSpeakingChanged", monitor.listener);
    participantSpeakerMonitors.delete(participantKey);

    return queueSpeakerEvent({
        type: "PARTICIPANT_LEFT",
        timestamp: new Date().toISOString(),
        participantId: participantKey,
        participantName: participant.displayName || "Unknown participant"
    });
}

function attachToExistingRemoteParticipants() {
    if (!call) return;

    const participants = Array.from(call.remoteParticipants || []);
    if (participants.length > 0) hasSeenRemoteParticipant = true;
    participants.forEach(monitorRemoteParticipant);
    scheduleAutoLeaveCheck("existing_remote_participants_attached");
}

async function endSpeakerTrackingSession(reason = "meeting_ended") {
    if (!speakerSessionId || speakerSessionEnded) {
        return;
    }

    speakerSessionEnded = true;

    for (const monitor of participantSpeakerMonitors.values()) {
        const participant = monitor.participant;

        if (monitor.isSpeaking) {
            monitor.isSpeaking = false;
            queueSpeakerEvent({
                type: "SPEAKER_STOPPED",
                timestamp: new Date().toISOString(),
                participantId: getParticipantKey(participant),
                participantName: participant.displayName || "Unknown participant",
                reason
            });
        }

        participant.off("isSpeakingChanged", monitor.listener);
    }

    participantSpeakerMonitors.clear();

    queueSpeakerEvent({
        type: "SESSION_ENDED",
        timestamp: new Date().toISOString(),
        reason
    });

    await speakerEventQueue;
    console.log("Speaker tracking session saved:", speakerSessionId);
}


// --------------------------------------------------
// UI ELEMENTS
// --------------------------------------------------

const teamsLinkInput =
    document.getElementById("teams-link");

const joinButton =
    document.getElementById("join-button");

const leaveButton =
    document.getElementById("leave-button");

const statusElement =
    document.getElementById("status");

installAutoLeaveSetting();


// --------------------------------------------------
// STATUS
// --------------------------------------------------

function setStatus(message) {

    console.log(message);

    if (statusElement) {
        statusElement.textContent = "Status: " + message;
    }
}


// --------------------------------------------------
// INITIALIZE ACS
// --------------------------------------------------

async function initialize() {

    try {

        setStatus(
            "Getting ACS token..."
        );

        const response =
            await fetch("/api/token");

        if (!response.ok) {

            throw new Error(
                "Failed to get ACS token"
            );
        }

        const data =
            await response.json();

        console.log(
            "ACS identity:",
            data.userId
        );

        tokenCredential =
            new AzureCommunicationTokenCredential(
                data.token
            );

        callClient =
            new CallClient();

        callAgent =
            await callClient.createCallAgent(
                tokenCredential,
                {
                    displayName:
                        "ScryAI Teams Bot"
                }
            );

        setStatus(
            "ACS client ready."
        );

        joinButton.disabled = false;

    } catch (error) {

        console.error(
            "ACS initialization failed:",
            error
        );

        setStatus(
            "Initialization failed: " +
            error.message
        );
    }
}


// --------------------------------------------------
// JOIN TEAMS MEETING
// --------------------------------------------------

async function joinMeeting() {

    const meetingLink =
        teamsLinkInput.value.trim();

    if (!meetingLink) {

        setStatus(
            "Please enter a Teams meeting URL."
        );

        return;
    }

    try {

        joinButton.disabled = true;

        setStatus(
            "Joining Teams meeting..."
        );

        // Create/resume AudioContext during the Join button gesture. Raw remote
        // audio is attached only after ACS reports the call as Connected.
        try {
            await prepareAudioContext();
        } catch (audioContextError) {
            console.warn("AudioContext preflight could not start yet:", audioContextError);
        }

        startSpeakerTrackingSession();

        // Join as a receive-only bot.
// Keep outgoing microphone muted to prevent echo.
        call =
            await callAgent.join(
                {
                    meetingLink: meetingLink
                },
                {
                    audioOptions: {
                        muted: true
                    }
                }
            );

        console.log(
            "Bot joined with outgoing microphone muted:",
            call.isMuted
        );

        call.on("isMutedChanged", () => {
            console.log(
                "Bot microphone mute state:",
                call.isMuted
            );
        });

        // Subscribe to participants joining/leaving, then attach to the current roster.
        call.on("remoteParticipantsUpdated", (event) => {
            event.added.forEach(participant => {
                hasSeenRemoteParticipant = true;
                monitorRemoteParticipant(participant);
            });
            event.removed.forEach(participant => {
                stopMonitoringRemoteParticipant(participant, "participant_left");
            });

            const currentCount = call ? call.remoteParticipants.length : 0;
            if (currentCount > 0) hasSeenRemoteParticipant = true;
            console.log("Remote participant roster updated:", {
                added: event.added.length,
                removed: event.removed.length,
                currentCount,
                autoLeaveEnabled: autoLeaveWhenEmpty
            });
            scheduleAutoLeaveCheck("remote_participants_updated");
        });

        attachToExistingRemoteParticipants();

        remoteAudioStreamsUpdatedHandler = (event) => {
            (event.added || []).forEach((stream) => {
                void attachRemoteAudioStream(stream);
            });
            (event.removed || []).forEach(removeRemoteAudioStream);
            console.log("ACS remote audio stream update:", {
                added: (event.added || []).length,
                removed: (event.removed || []).length,
                total: call ? call.remoteAudioStreams.length : 0
            });
        };
        call.on("remoteAudioStreamsUpdated", remoteAudioStreamsUpdatedHandler);

        console.log(
            "Call object:",
            call
        );

        console.log(
            "Connected / joining meeting."
        );


        // ------------------------------------------
        // CALL STATE
        // ------------------------------------------

        call.on("stateChanged", async () => {
            await handleCallStateChange();

            if (call && call.state === "Connected") {
                console.log("================================");
                console.log("TEAMS MEETING CONNECTED");
                console.log("Remote participants:", call.remoteParticipants.length);
                console.log("Remote audio streams:", call.remoteAudioStreams.length);
                console.log("================================");

                attachToExistingRemoteParticipants();
                console.log("Speaker tracking active. Checking ACS remote audio streams...");
                void connectExistingRemoteAudioStreams();
            }
        });

        // Read the current state immediately after subscribing so an early
        // InLobby/Connected transition is not missed.
        await handleCallStateChange();
        if (call && call.state === "Connected") {
            await connectExistingRemoteAudioStreams();
        }

        leaveButton.disabled = false;

        setStatus(
            "Connected / joining meeting."
        );


    } catch (error) {

        await stopLiveAudioCapture("join_failed");
        const callEndReason = getCallEndReasonDetails(call);
        await publishCallStatus("failed", {
            inLobbyAtUtc,
            callConnectedAtUtc,
            meetingEndedAtUtc,
            message: error.message || "ACS join request failed.",
            errorName: error.name || null,
            errorCode: error.code || null,
            callEndReason
        });
        await endSpeakerTrackingSession("join_failed");

        console.error(
            "Join failed:",
            error
        );

        setStatus(
            "Join failed: " +
            error.message
        );

        joinButton.disabled = false;
        leaveButton.disabled = true;
    }
}


// --------------------------------------------------
// LIVE PCM AUDIO STREAMING
// --------------------------------------------------

function getAudioWebSocketUrl() {
    const protocol = window.location.protocol === "https:" ? "wss:" : "ws:";
    return `${protocol}//${window.location.host}/ws/audio`;
}

function waitForWebSocketOpen(socket, timeoutMs = 10000) {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("Timed out opening audio WebSocket.")), timeoutMs);
        const onOpen = () => {
            clearTimeout(timer);
            socket.removeEventListener("error", onError);
            resolve();
        };
        const onError = () => {
            clearTimeout(timer);
            socket.removeEventListener("open", onOpen);
            reject(new Error("Could not open audio WebSocket. Check that the Node server and /ws/audio endpoint are running."));
        };
        socket.addEventListener("open", onOpen, { once: true });
        socket.addEventListener("error", onError, { once: true });
    });
}

function waitForAudioSocketMessage(socket, expectedType, timeoutMs = 30000) {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
            socket.removeEventListener("message", onMessage);
            reject(new Error(`Timed out waiting for audio server message: ${expectedType}`));
        }, timeoutMs);

        const onMessage = (event) => {
            let message;
            try {
                message = JSON.parse(event.data);
            } catch {
                return;
            }

            if (message.type === "error") {
                clearTimeout(timer);
                socket.removeEventListener("message", onMessage);
                reject(new Error(message.message || "Audio WebSocket server error."));
                return;
            }

            if (message.type === expectedType) {
                clearTimeout(timer);
                socket.removeEventListener("message", onMessage);
                resolve(message);
            }
        };

        socket.addEventListener("message", onMessage);
    });
}

function sendAudioControl(payload) {
    if (audioSocket && audioSocket.readyState === WebSocket.OPEN) {
        audioSocket.send(JSON.stringify(payload));
    }
}

function attachAudioContextStateMonitor(context) {
    if (audioContextStateHandler) {
        try { context.removeEventListener("statechange", audioContextStateHandler); } catch { }
    }

    audioContextLastState = context.state;
    audioContextStateHandler = () => {
        const state = context.state;
        const nowPerf = performance.now();
        const nowUtc = new Date().toISOString();
        const previousState = audioContextLastState;
        audioContextLastState = state;

        const eventRecord = {
            state,
            previousState,
            timestampUtc: nowUtc,
            performanceTimeMs: Number(nowPerf.toFixed(3)),
            audioContextTimeSeconds: Number(context.currentTime.toFixed(6))
        };
        if (callConnectedPerfMs !== null) {
            audioContextEvents.push(eventRecord);
            sendAudioControl({ type: "contextState", ...eventRecord });
        }

        if (callConnectedPerfMs !== null && state !== "running") {
            if (audioContextSuspendedAtPerfMs === null) {
                audioContextSuspendedAtPerfMs = nowPerf;
                audioContextSuspendedAtUtc = nowUtc;
            }
        } else if (callConnectedPerfMs !== null && state === "running" && audioContextSuspendedAtPerfMs !== null) {
            const monotonicDurationMs = Math.max(0, nowPerf - audioContextSuspendedAtPerfMs);
            const wallDurationMs = Math.max(0, Date.parse(nowUtc) - Date.parse(audioContextSuspendedAtUtc));
            // UTC is the anchor used for the WAV timeline, so reconcile suspension
            // samples against that same wall-clock interval. Keep monotonic time too
            // to expose any system-clock adjustment in the metadata.
            const durationMs = wallDurationMs;
            const sampleCount = Math.max(0, Math.round(durationMs * AUDIO_SAMPLE_RATE / 1000));
            const gap = {
                reason: "audio_context_suspended",
                startSample: audioLastSampleIndexSent,
                sampleCount,
                startUtc: audioContextSuspendedAtUtc,
                endUtc: nowUtc,
                durationMs: Number(durationMs.toFixed(3)),
                monotonicDurationMs: Number(monotonicDurationMs.toFixed(3)),
                wallVsMonotonicDeltaMs: Number((wallDurationMs - monotonicDurationMs).toFixed(3))
            };

            if (sampleCount > 0 && audioWorkletNode) {
                // The worklet inserts this interval into the same sample timeline.
                audioWorkletNode.port.postMessage({ type: "insertSilence", ...gap });
                console.warn("[AUDIO CLOCK] Inserted silence for AudioContext interruption:", gap);
            } else if (sampleCount > 0) {
                // Before the processor exists, Connected-to-processor startup padding
                // already covers this period; retain the event for diagnostics only.
                console.warn("[AUDIO CLOCK] Context resumed before worklet startup; startup padding covers this interval:", gap);
            }

            audioContextSuspendedAtPerfMs = null;
            audioContextSuspendedAtUtc = null;
        }
    };

    context.addEventListener("statechange", audioContextStateHandler);
}

async function prepareAudioContext() {
    if (audioContext) {
        if (audioContext.state !== "running" && callConnectedPerfMs !== null && audioContextSuspendedAtPerfMs === null) {
            // The context may have suspended while the bot was waiting in the lobby.
            // Start measuring only from Connected, never from InLobby.
            audioContextSuspendedAtPerfMs = performance.now();
            audioContextSuspendedAtUtc = new Date().toISOString();
            audioContextEvents.push({
                state: "suspended_at_connected_capture_start",
                previousState: audioContextLastState,
                timestampUtc: audioContextSuspendedAtUtc,
                performanceTimeMs: Number(audioContextSuspendedAtPerfMs.toFixed(3)),
                audioContextTimeSeconds: Number(audioContext.currentTime.toFixed(6))
            });
        }
        if (audioContext.state === "suspended") await audioContext.resume();
        attachAudioContextStateMonitor(audioContext);
        return audioContext;
    }

    const AudioContextConstructor = window.AudioContext || window.webkitAudioContext;
    if (!AudioContextConstructor) {
        throw new Error("This browser does not support Web Audio API.");
    }

    try {
        audioContext = new AudioContextConstructor({ sampleRate: AUDIO_SAMPLE_RATE });
    } catch {
        // The worklet resamples if the browser cannot create a 16 kHz context.
        audioContext = new AudioContextConstructor();
    }

    if (audioContext.state !== "running" && callConnectedPerfMs !== null && audioContextSuspendedAtPerfMs === null) {
        audioContextSuspendedAtPerfMs = performance.now();
        audioContextSuspendedAtUtc = new Date().toISOString();
        audioContextEvents.push({
            state: "suspended_at_connected_capture_start",
            previousState: audioContextLastState,
            timestampUtc: audioContextSuspendedAtUtc,
            performanceTimeMs: Number(audioContextSuspendedAtPerfMs.toFixed(3)),
            audioContextTimeSeconds: Number(audioContext.currentTime.toFixed(6))
        });
    }
    if (audioContext.state === "suspended") await audioContext.resume();
    attachAudioContextStateMonitor(audioContext);
    console.log("AudioContext ready:", audioContext.sampleRate, "Hz");
    return audioContext;
}

function encodePcmChunkFrame(payload) {
    const pcmBytes = new Uint8Array(payload.buffer);
    const frame = new ArrayBuffer(AUDIO_FRAME_HEADER_BYTES + pcmBytes.byteLength);
    const view = new DataView(frame);
    view.setUint32(0, AUDIO_FRAME_MAGIC, true);
    view.setUint32(4, payload.sequence >>> 0, true);
    view.setBigUint64(8, BigInt(payload.sampleOffset), true);
    view.setUint32(16, payload.sampleCount >>> 0, true);
    new Uint8Array(frame, AUDIO_FRAME_HEADER_BYTES).set(pcmBytes);
    return frame;
}

async function startPcmCapturePipeline() {
    if (audioWorkletNode && audioSocket && audioSocket.readyState === WebSocket.OPEN) {
        return;
    }
    if (audioStartPromise) return audioStartPromise;

    audioStartPromise = (async () => {
        if (!speakerSessionId) throw new Error("No active meeting session ID for audio capture.");

        const socket = new WebSocket(getAudioWebSocketUrl());
        socket.binaryType = "arraybuffer";
        audioSocket = socket;
        socket.addEventListener("close", (event) => {
            console.warn("Audio WebSocket closed:", event.code, event.reason || "");
        });
        socket.addEventListener("error", (event) => {
            console.error("Audio WebSocket error:", event);
        });

        await waitForWebSocketOpen(socket);
        const startedAck = waitForAudioSocketMessage(socket, "started", 15000);
        captureRequestedAtUtc = new Date().toISOString();
        socket.send(JSON.stringify({
            type: "start",
            protocolVersion: 2,
            sessionId: speakerSessionId,
            sampleRate: AUDIO_SAMPLE_RATE,
            channels: 1,
            bitDepth: 16,
            inLobbyAtUtc,
            meetingConnectedAtUtc: callConnectedAtUtc,
            captureRequestedAtUtc,
            audioTimelineStartAtUtc: callConnectedAtUtc
        }));
        await startedAck;

        const context = await prepareAudioContext();
        await context.audioWorklet.addModule("/pcm-capture-worklet.js");

        // Pad the small async setup interval so WAV sample zero maps to the
        // first ACS Connected timestamp, not to the later remote-track arrival.
        const connectedEpochMs = callConnectedAtUtc ? Date.parse(callConnectedAtUtc) : NaN;
        const processorSetupAtUtc = new Date().toISOString();
        audioInitialSilenceSamples = Number.isFinite(connectedEpochMs)
            ? Math.max(0, Math.round((Date.parse(processorSetupAtUtc) - connectedEpochMs) * AUDIO_SAMPLE_RATE / 1000))
            : 0;

        audioWorkletNode = new AudioWorkletNode(context, "pcm16-capture", {
            numberOfInputs: 1,
            numberOfOutputs: 1,
            outputChannelCount: [1],
            channelCount: 2,
            channelCountMode: "max",
            channelInterpretation: "speakers",
            processorOptions: {
                initialSilenceSamples: audioInitialSilenceSamples,
                connectedAtUtc: callConnectedAtUtc,
                processorSetupAtUtc
            }
        });

        audioWorkletNode.port.onmessage = (event) => {
            const payload = event.data || {};
            if (payload.type === "chunk" && payload.buffer) {
                if (audioSocket && audioSocket.readyState === WebSocket.OPEN) {
                    const frame = encodePcmChunkFrame(payload);
                    audioSocket.send(frame);
                    audioChunkCount += 1;
                    audioBytesSent += payload.buffer.byteLength;
                    audioSamplesSent += payload.sampleCount;
                    audioLastSampleIndexSent = Math.max(
                        audioLastSampleIndexSent,
                        payload.sampleOffset + payload.sampleCount
                    );
                    if (audioSocket.bufferedAmount > 8 * 1024 * 1024) {
                        console.warn("[LIVE AUDIO] WebSocket send queue is growing:", audioSocket.bufferedAmount);
                    }
                    if (audioChunkCount === 1 || audioChunkCount % 10 === 0) {
                        console.log(`[LIVE AUDIO] Sent chunk=${audioChunkCount} sequence=${payload.sequence} firstSample=${payload.sampleOffset} samples=${payload.sampleCount} totalSamples=${audioLastSampleIndexSent}`);
                    }
                } else {
                    console.error("PCM chunk could not be sent because the audio WebSocket is not open:", {
                        sequence: payload.sequence,
                        sampleOffset: payload.sampleOffset,
                        sampleCount: payload.sampleCount
                    });
                }
            } else if (payload.type === "gap") {
                const gap = {
                    reason: payload.reason || "audio_gap",
                    startSample: payload.startSample,
                    sampleCount: payload.sampleCount,
                    startUtc: payload.startUtc || null,
                    endUtc: payload.endUtc || null,
                    durationMs: payload.durationMs || null
                };
                audioTimelineGaps.push(gap);
                sendAudioControl({ type: "gap", ...gap });
                console.warn("[LIVE AUDIO] Timeline gap recorded:", gap);
            } else if (payload.type === "stats") {
                audioWorkletStats.push({ ...payload, timestampUtc: new Date().toISOString() });
                console.log("[LIVE AUDIO] AudioWorklet input stats:", payload);
                sendAudioControl({ type: "stats", stats: payload, timestampUtc: new Date().toISOString() });
            } else if (payload.type === "processorStarted") {
                audioProcessorStartedAtUtc = new Date().toISOString();
                console.log("[LIVE AUDIO] AudioWorklet processor started:", {
                    audioProcessorStartedAtUtc,
                    initialSilenceSamples: audioInitialSilenceSamples,
                    initialSilenceMs: Number((audioInitialSilenceSamples / AUDIO_SAMPLE_RATE * 1000).toFixed(2)),
                    contextSampleRate: payload.contextSampleRate
                });
                sendAudioControl({
                    type: "timeline",
                    audioCapturePipelineReadyAtUtc,
                    audioProcessorStartedAtUtc,
                    audioInitialSilenceSamples,
                    audioTimelineStartAtUtc: callConnectedAtUtc,
                    contextSampleRate: payload.contextSampleRate
                });
            } else if (payload.type === "flushed" && audioFlushResolver) {
                audioWorkletTotalSampleCount = payload.totalSampleCount;
                const resolve = audioFlushResolver;
                audioFlushResolver = null;
                resolve();
            }
        };

        audioSilentGain = context.createGain();
        audioSilentGain.gain.value = 0;
        audioWorkletNode.connect(audioSilentGain);
        audioSilentGain.connect(context.destination);
        audioCapturePipelineReadyAtUtc = new Date().toISOString();

        audioChunkCount = 0;
        audioBytesSent = 0;
        audioSamplesSent = 0;
        audioLastSampleIndexSent = 0;
        audioWorkletTotalSampleCount = null;
        console.log("[LIVE AUDIO] WebSocket session started; sample-indexed PCM16 mono 16 kHz capture is ready.");
    })();

    try {
        await audioStartPromise;
    } catch (error) {
        if (audioSocket && audioSocket.readyState === WebSocket.OPEN) {
            try { audioSocket.send(JSON.stringify({ type: "end", reason: "capture_start_failed" })); } catch { }
            audioSocket.close();
        }
        audioSocket = null;
        throw error;
    } finally {
        audioStartPromise = null;
    }
}

async function attachRemoteAudioStream(remoteAudioStream) {
    if (!remoteAudioStream || audioSourcesByRemoteStream.has(remoteAudioStream)) return;
    if (audioStreamAttachPromises.has(remoteAudioStream)) return audioStreamAttachPromises.get(remoteAudioStream);
    if (!call || call.state !== "Connected") return;

    const attachPromise = (async () => {
        console.log("[LIVE AUDIO] Requesting raw MediaStream from ACS remote audio stream...");
        const mediaStream = await remoteAudioStream.getMediaStream();
        if (!mediaStream || typeof mediaStream.getAudioTracks !== "function") {
            throw new Error("ACS getMediaStream() did not return a browser MediaStream.");
        }

        const tracks = mediaStream.getAudioTracks();
        if (!tracks.length) throw new Error("ACS remote MediaStream contains no audio tracks.");

        await startPcmCapturePipeline();
        const source = audioContext.createMediaStreamSource(mediaStream);
        source.connect(audioWorkletNode);
        audioSourcesByRemoteStream.set(remoteAudioStream, { source, mediaStream });

        console.log("[LIVE AUDIO] Remote stream connected to AudioWorklet:", {
            remoteAudioStreamCount: call.remoteAudioStreams.length,
            trackCount: tracks.length,
            trackStates: tracks.map(track => ({
                enabled: track.enabled,
                muted: track.muted,
                readyState: track.readyState,
                settings: typeof track.getSettings === "function" ? track.getSettings() : null
            })),
            audioContextSampleRate: audioContext.sampleRate
        });
        setStatus("Connected — live meeting audio streaming to backend (PCM 16 kHz mono).");
    })();

    audioStreamAttachPromises.set(remoteAudioStream, attachPromise);
    try {
        await attachPromise;
    } catch (error) {
        console.error("[LIVE AUDIO] Could not attach ACS remote audio stream:", error);
        const message = String(error?.message || error);
        setStatus(`Connected, but raw audio capture failed: ${message}`);
    } finally {
        audioStreamAttachPromises.delete(remoteAudioStream);
    }
}

function removeRemoteAudioStream(remoteAudioStream) {
    const entry = audioSourcesByRemoteStream.get(remoteAudioStream);
    if (!entry) return;
    try { entry.source.disconnect(); } catch { }
    audioSourcesByRemoteStream.delete(remoteAudioStream);
    console.log("[LIVE AUDIO] Remote audio stream removed from capture graph.");
}

async function connectExistingRemoteAudioStreams() {
    if (!call || call.state !== "Connected") return;
    const streams = Array.from(call.remoteAudioStreams || []);
    console.log("[LIVE AUDIO] Current ACS remote audio stream count:", streams.length);
    if (streams.length === 0) {
        console.log("[LIVE AUDIO] Waiting for call.remoteAudioStreamsUpdated...");
        return;
    }
    await Promise.all(streams.map(stream => attachRemoteAudioStream(stream)));
}

async function stopLiveAudioCapture(reason = "meeting_ended") {
    if (audioStopPromise) return audioStopPromise;

    audioStopPromise = (async () => {
        // A remote stream can be attaching at the same moment the call ends.
        // Let that setup settle before deciding whether there is a WS session to finalize.
        if (audioStartPromise) {
            try { await audioStartPromise; } catch (error) {
                console.warn("Audio capture was still starting while shutdown began:", error);
            }
        }

        const hasAudioSession = Boolean(audioSocket || audioWorkletNode || audioContext);
        if (!hasAudioSession) return;

        console.log(`[LIVE AUDIO] Stopping capture: ${reason}`);

        // Ask the worklet to emit its final partial PCM chunk before ending the WS session.
        if (audioWorkletNode && audioSocket && audioSocket.readyState === WebSocket.OPEN) {
            const flushPromise = new Promise((resolve, reject) => {
                const timer = setTimeout(() => {
                    audioFlushResolver = null;
                    reject(new Error("Timed out flushing the final PCM chunk from AudioWorklet."));
                }, 5000);
                audioFlushResolver = () => {
                    clearTimeout(timer);
                    resolve();
                };
            });
            audioWorkletNode.port.postMessage({ type: "flush" });
            try {
                await flushPromise;
            } catch (error) {
                console.warn("AudioWorklet flush warning:", error);
            }
        }

        if (audioSocket && audioSocket.readyState === WebSocket.OPEN) {
            const finalizedPromise = waitForAudioSocketMessage(audioSocket, "finalized", 30000);
            meetingEndedAtUtc = meetingEndedAtUtc || new Date().toISOString();
            if (meetingEndedPerfMs === null) meetingEndedPerfMs = performance.now();
            const connectedEpochMs = callConnectedAtUtc ? Date.parse(callConnectedAtUtc) : NaN;
            const endedEpochMs = Date.parse(meetingEndedAtUtc);
            const utcElapsedMs = Number.isFinite(connectedEpochMs) && Number.isFinite(endedEpochMs)
                ? Math.max(0, endedEpochMs - connectedEpochMs)
                : null;
            const monotonicElapsedMs = callConnectedPerfMs === null
                ? null
                : Math.max(0, meetingEndedPerfMs - callConnectedPerfMs);
            const targetSampleCount = utcElapsedMs === null
                ? null
                : Math.max(0, Math.round(utcElapsedMs * AUDIO_SAMPLE_RATE / 1000));
            const pendingSuspension = audioContextSuspendedAtPerfMs === null
                ? null
                : {
                    startUtc: audioContextSuspendedAtUtc,
                    endUtc: meetingEndedAtUtc,
                    durationMs: Number(Math.max(0, Date.parse(meetingEndedAtUtc) - Date.parse(audioContextSuspendedAtUtc)).toFixed(3)),
                    monotonicDurationMs: Number(Math.max(0, meetingEndedPerfMs - audioContextSuspendedAtPerfMs).toFixed(3)),
                    reason: "audio_context_still_suspended_at_end"
                };
            if (pendingSuspension) audioContextEvents.push({ state: "still_suspended_at_end", ...pendingSuspension });

            audioSocket.send(JSON.stringify({
                type: "end",
                protocolVersion: 2,
                reason,
                meetingEndedAtUtc,
                meetingEndedPerfMs,
                inLobbyAtUtc,
                meetingConnectedAtUtc: callConnectedAtUtc,
                captureRequestedAtUtc,
                audioCapturePipelineReadyAtUtc,
                audioProcessorStartedAtUtc,
                audioInitialSilenceSamples,
                targetSampleCount,
                utcElapsedMs,
                monotonicElapsedMs,
                wallClockVsMonotonicDeltaMs: utcElapsedMs === null || monotonicElapsedMs === null
                    ? null
                    : Number((utcElapsedMs - monotonicElapsedMs).toFixed(3)),
                workletReportedSampleCount: audioWorkletTotalSampleCount,
                audioLastSampleIndexSent,
                audioSamplesSent,
                contextStateEvents: audioContextEvents,
                timelineGaps: audioTimelineGaps,
                workletStats: audioWorkletStats,
                pendingSuspension
            }));
            try {
                const result = await finalizedPromise;
                console.log("[LIVE AUDIO] FINAL WAV SAVED:", result.wavPath || "(no audio samples)", {
                    chunks: result.chunks,
                    bytes: result.audioBytes,
                    durationSeconds: result.durationSeconds
                });
                if (result.wavPath) setStatus(`Audio finalized: ${result.wavPath}`);
            } catch (error) {
                console.error("[LIVE AUDIO] Server did not confirm WAV finalization:", error);
            }
        }

        for (const { source } of audioSourcesByRemoteStream.values()) {
            try { source.disconnect(); } catch { }
        }
        audioSourcesByRemoteStream.clear();
        audioStreamAttachPromises.clear();

        if (audioWorkletNode) {
            try { audioWorkletNode.port.onmessage = null; audioWorkletNode.disconnect(); } catch { }
        }
        if (audioSilentGain) {
            try { audioSilentGain.disconnect(); } catch { }
        }
        if (audioContext) {
            if (audioContextStateHandler) {
                try { audioContext.removeEventListener("statechange", audioContextStateHandler); } catch { }
            }
            try { await audioContext.close(); } catch (error) {
                console.warn("AudioContext close warning:", error);
            }
        }
        audioContextStateHandler = null;
        if (audioSocket && audioSocket.readyState < WebSocket.CLOSING) {
            try { audioSocket.close(1000, "capture complete"); } catch { }
        }

        audioSocket = null;
        audioWorkletNode = null;
        audioSilentGain = null;
        audioContext = null;
        audioFlushResolver = null;
    })();

    try {
        await audioStopPromise;
    } finally {
        audioStopPromise = null;
    }
}


// --------------------------------------------------
// LEAVE MEETING
// --------------------------------------------------

async function leaveMeeting(reason = "bot_left_meeting") {
    if (leaveInProgress) return;
    leaveInProgress = true;

    try {
        meetingEndedAtUtc = meetingEndedAtUtc || new Date().toISOString();
        if (meetingEndedPerfMs === null) meetingEndedPerfMs = performance.now();

        // Flush audio to the Node server before hanging up the ACS call.
        await stopLiveAudioCapture(reason);

        if (call && call.state !== "Disconnected") {
            await publishCallStatus("disconnected", {
                callState: call.state,
                reason,
                inLobbyAtUtc,
                callConnectedAtUtc,
                meetingEndedAtUtc,
                message: reason === "all_remote_participants_left"
                    ? "All remote participants left; bot automatically left the meeting."
                    : "Bot left the Teams meeting on request."
            });
            terminalCallStatusPublished = true;
        }

        await endSpeakerTrackingSession(reason);

        if (call) {
            const callToLeave = call;
            if (remoteAudioStreamsUpdatedHandler) {
                try { callToLeave.off("remoteAudioStreamsUpdated", remoteAudioStreamsUpdatedHandler); } catch { }
            }
            remoteAudioStreamsUpdatedHandler = null;
            await callToLeave.hangUp();
            if (call === callToLeave) call = null;
        }

        joinButton.disabled = false;
        leaveButton.disabled = true;
        setStatus(reason === "all_remote_participants_left"
            ? "All participants left. Bot has left the meeting and audio is finalized."
            : "Left Teams meeting.");
    } catch (error) {
        console.error("Leave failed:", error);
        setStatus("Leave failed: " + error.message);
    } finally {
        leaveInProgress = false;
    }
}


// --------------------------------------------------
// BUTTON EVENTS
// --------------------------------------------------

joinButton.addEventListener(
    "click",
    joinMeeting
);


leaveButton.addEventListener(
    "click",
    () => leaveMeeting("bot_left_meeting")
);


// --------------------------------------------------
// INITIAL STATE
// --------------------------------------------------

joinButton.disabled =
    true;

leaveButton.disabled =
    true;


// --------------------------------------------------
// START ACS
// --------------------------------------------------

initialize();