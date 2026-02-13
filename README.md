# local-ai-clients

Client applications for [local-ai-server](../local-ai-server/). Voice transcription GUI, pronunciation drill, and headless test harness.

## Clients

### voice-test-gui

Interactive GUI for real-time voice transcription with waveform visualization, stability detection, pronunciation drill mode, LLM chat, and TTS.

```batch
clients\voice-test-gui\build.bat
bin\voice-test-gui.exe
```

Hold Space to talk. Press D for pronunciation drill mode.

### voice-test-headless

Offline test harness for comparing transcription approaches against recorded WAV files.

```batch
clients\voice-test-headless\build.bat
bin\voice-test-headless.exe --mode sim --interval 3 recording.wav
```

Modes: `sim` (full pipeline simulation), `retranscribe`, `timestamps`, `ts-sweep`.

## Architecture

```
voice-test-gui.exe ──HTTP──> local-ai-server:8090 (ASR)
                   ──HTTP──> llama-server:8042    (LLM, optional)
                   ──SAPI──> Windows TTS          (built-in)
```

Clients are thin: all AI inference runs on the server. Clients handle audio capture, UI, and HTTP communication.

## Prerequisites

1. Visual Studio with C++ workload (MSVC compiler)
2. [local-ai-server](../local-ai-server/) running for ASR

No external dependencies (DEPS_ROOT not required). TTS uses Windows SAPI.

## Repository Structure

```
local-ai-clients/
├── clients/
│   ├── voice-test-gui/        GUI with waveform, drill, LLM chat
│   │   ├── build.bat
│   │   └── src/
│   └── voice-test-headless/   Offline test harness
│       ├── build.bat
│       └── src/
├── shared/                    Shared HTTP client library
│   ├── asr_client.h
│   └── asr_client.c
└── data/                      Drill sentence banks
    └── drill_sentences.txt
```

## Platform

Windows only (Media Foundation for audio capture, GDI for rendering, WinHTTP for server communication).
