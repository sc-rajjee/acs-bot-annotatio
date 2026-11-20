/*
 * ACS remote audio -> mono PCM16 @ 16 kHz.
 *
 * Protocol v2 guarantees every emitted PCM block carries:
 *   sequence, firstSampleIndex, sampleCount.
 *
 * The AudioWorklet owns the sample counter. Known AudioContext suspension
 * intervals are inserted as explicit silence ranges, while normal microphone/
 * remote-stream silence remains ordinary audio silence.
 */
class Pcm16CaptureProcessor extends AudioWorkletProcessor {
    constructor(options = {}) {
        super();

        const processorOptions = options.processorOptions || {};
        this.targetSampleRate = 16000;
        this.samplesPerChunk = 16000;
        this.chunk = new Int16Array(this.samplesPerChunk);
        this.chunkOffset = 0;
        this.chunkStartSample = 0;
        this.sampleCursor = 0;
        this.sequence = 0;
        this.phase = 0;
        this.previousFiltered = 0;
        this.hasPreviousFiltered = false;
        this.stopped = false;
        this.processorStartedReported = false;

        this.reportFrames = 0;
        this.reportEmptyInputBlocks = 0;
        this.reportEmptyInputFrames = 0;
        this.reportSilentInputFrames = 0;
        this.reportNonSilentInputFrames = 0;
        this.reportInputBlocks = 0;

        const cutoffHz = 7000;
        const omega = 2 * Math.PI * cutoffHz / sampleRate;
        const cosine = Math.cos(omega);
        const alpha = Math.sin(omega) / (2 * Math.SQRT1_2);
        const a0 = 1 + alpha;
        this.b0 = ((1 - cosine) / 2) / a0;
        this.b1 = (1 - cosine) / a0;
        this.b2 = this.b0;
        this.a1 = (-2 * cosine) / a0;
        this.a2 = (1 - alpha) / a0;
        this.x1 = 0;
        this.x2 = 0;
        this.y1 = 0;
        this.y2 = 0;

        this.port.onmessage = (event) => {
            const message = event.data || {};

            if (message.type === "insertSilence") {
                const count = Number(message.sampleCount);
                if (Number.isSafeInteger(count) && count > 0) {
                    // Keep chunk/sample offsets contiguous around the known gap.
                    this.emitChunk();
                    const startSample = this.sampleCursor;
                    this.port.postMessage({
                        type: "gap",
                        startSample,
                        sampleCount: count,
                        reason: message.reason || "audio_context_suspended",
                        startUtc: message.startUtc || null,
                        endUtc: message.endUtc || null,
                        durationMs: Number(message.durationMs) || null
                    });
                    this.pushSilence(count);
                    this.resetFilter();
                }
                return;
            }

            if (message.type === "flush") {
                this.emitChunk();
                this.port.postMessage({
                    type: "flushed",
                    totalSampleCount: this.sampleCursor,
                    nextSequence: this.sequence
                });
                this.stopped = true;
            }
        };

        const initialSilenceSamples = Number(processorOptions.initialSilenceSamples) || 0;
        if (Number.isSafeInteger(initialSilenceSamples) && initialSilenceSamples > 0) {
            this.port.postMessage({
                type: "gap",
                startSample: 0,
                sampleCount: initialSilenceSamples,
                reason: "connected_to_audio_processor_startup",
                startUtc: processorOptions.connectedAtUtc || null,
                endUtc: processorOptions.processorSetupAtUtc || null,
                durationMs: initialSilenceSamples * 1000 / this.targetSampleRate
            });
            this.pushSilence(initialSilenceSamples);
        }
    }

    resetFilter() {
        this.phase = 0;
        this.previousFiltered = 0;
        this.hasPreviousFiltered = false;
        this.x1 = this.x2 = this.y1 = this.y2 = 0;
    }

    lowPass(value) {
        const output = this.b0 * value
            + this.b1 * this.x1
            + this.b2 * this.x2
            - this.a1 * this.y1
            - this.a2 * this.y2;

        this.x2 = this.x1;
        this.x1 = value;
        this.y2 = this.y1;
        this.y1 = output;
        return output;
    }

    emitChunk() {
        if (this.chunkOffset <= 0) return;

        const completed = this.chunk.slice(0, this.chunkOffset);
        const message = {
            type: "chunk",
            buffer: completed.buffer,
            sequence: this.sequence++,
            sampleOffset: this.chunkStartSample,
            sampleCount: completed.length
        };
        this.port.postMessage(message, [completed.buffer]);

        this.chunkOffset = 0;
        this.chunkStartSample = this.sampleCursor;
    }

    pushOutputSample(value) {
        if (this.chunkOffset === 0) this.chunkStartSample = this.sampleCursor;

        const clipped = Math.max(-1, Math.min(1, value));
        const int16 = clipped < 0
            ? Math.round(clipped * 32768)
            : Math.round(clipped * 32767);

        this.chunk[this.chunkOffset++] = int16;
        this.sampleCursor += 1;

        if (this.chunkOffset === this.samplesPerChunk) {
            this.emitChunk();
        }
    }

    pushSilence(sampleCount) {
        // Fill output buffers efficiently while preserving exact sample offsets.
        for (let index = 0; index < sampleCount; index++) {
            this.pushOutputSample(0);
        }
    }

    reportStatsIfDue(frameCount) {
        this.reportFrames += frameCount;
        if (this.reportFrames < sampleRate * 5) return;

        this.port.postMessage({
            type: "stats",
            contextSampleRate: sampleRate,
            targetSampleRate: this.targetSampleRate,
            inputBlocks: this.reportInputBlocks,
            emptyInputBlocks: this.reportEmptyInputBlocks,
            emptyInputFrames: this.reportEmptyInputFrames,
            silentInputFrames: this.reportSilentInputFrames,
            nonSilentInputFrames: this.reportNonSilentInputFrames,
            outputSampleCursor: this.sampleCursor,
            nextSequence: this.sequence
        });

        this.reportFrames = 0;
        this.reportInputBlocks = 0;
        this.reportEmptyInputBlocks = 0;
        this.reportEmptyInputFrames = 0;
        this.reportSilentInputFrames = 0;
        this.reportNonSilentInputFrames = 0;
    }

    process(inputs, outputs) {
        for (const output of outputs) {
            for (const channel of output) channel.fill(0);
        }
        if (this.stopped) return false;

        if (!this.processorStartedReported) {
            this.processorStartedReported = true;
            this.port.postMessage({
                type: "processorStarted",
                contextSampleRate: sampleRate,
                totalSampleCount: this.sampleCursor,
                nextSequence: this.sequence
            });
        }

        const channels = inputs[0] || [];
        const outputFrameCount = outputs[0]?.[0]?.length || 128;
        const hasInput = channels.length > 0
            && channels[0]
            && channels[0].length > 0;
        const frameCount = hasInput ? channels[0].length : outputFrameCount;
        const channelCount = hasInput ? channels.length : 0;

        this.reportInputBlocks += 1;
        if (!hasInput) {
            this.reportEmptyInputBlocks += 1;
            this.reportEmptyInputFrames += frameCount;
        }

        for (let frame = 0; frame < frameCount; frame++) {
            let mono = 0;

            if (channelCount > 0) {
                let includedChannels = 0;
                for (let channel = 0; channel < channelCount; channel++) {
                    const data = channels[channel];
                    if (data && frame < data.length) {
                        mono += data[frame];
                        includedChannels += 1;
                    }
                }
                if (includedChannels > 0) mono /= includedChannels;
            }

            if (Math.abs(mono) < 1e-5) this.reportSilentInputFrames += 1;
            else this.reportNonSilentInputFrames += 1;

            const filtered = this.lowPass(mono);
            const phaseBefore = this.phase;
            this.phase += this.targetSampleRate;

            if (this.phase >= sampleRate) {
                const fraction = sampleRate <= this.targetSampleRate
                    ? 1
                    : Math.max(0, Math.min(
                        1,
                        (sampleRate - phaseBefore) / this.targetSampleRate
                    ));
                const previous = this.hasPreviousFiltered
                    ? this.previousFiltered
                    : filtered;
                const interpolated = previous + (filtered - previous) * fraction;

                this.phase -= sampleRate;
                this.pushOutputSample(interpolated);
            }

            this.previousFiltered = filtered;
            this.hasPreviousFiltered = true;
        }

        this.reportStatsIfDue(frameCount);
        return true;
    }
}

registerProcessor("pcm16-capture", Pcm16CaptureProcessor);
