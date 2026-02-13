# CLAUDE.md

Guidance for Claude Code when working with this repository.

## Project Overview

Client applications for local-ai-server. Voice transcription GUI, pronunciation drill, and headless test tools. All AI inference happens on the server; clients handle audio capture, UI, and HTTP.

## Build

Each client has its own `build.bat` in its directory:

```batch
clients/voice-test-gui/build.bat       :: GUI with waveform, drill, LLM
clients/voice-test-headless/build.bat   :: Offline test harness
```

Run `.bat` files directly from Git Bash (no cmd /c wrapper). Output goes to `bin/` and `build/` at repo root.

MSVC environment is auto-detected. No external dependencies required (DEPS_ROOT not needed).

## Critical Rules

- **No emojis** in code, comments, commit messages, or documentation.
- **Running .bat files**: Run directly with forward slashes from Git Bash.
- **Zero warnings**: Fix all compiler warnings before committing.
- **No external dependencies** without approval. Current approved: Windows APIs, WinHTTP, SAPI.

## Structure

- `clients/voice-test-gui/src/main.c` -- Main GUI (~4000 lines, #includes drill.c and drill_render.c)
- `clients/voice-test-gui/src/drill.h/c` -- Pronunciation drill logic
- `clients/voice-test-gui/src/drill_render.c` -- GDI drill panel rendering
- `clients/voice-test-headless/src/main.c` -- Offline transcription test harness
- `shared/asr_client.h/c` -- HTTP client for local-ai-server (WinHTTP, SSE streaming)
- `data/drill_sentences.txt` -- Pipe-delimited HSK sentence bank

## Related Repositories

- [local-ai-server](https://github.com/rmtew/local-ai-server) -- ASR inference server (Qwen3-ASR)
- [lifeapp](https://github.com/rmtew/lifeapp) -- Original project (voice notes extracted from here)

## Key Patterns

- GUI uses `#include "drill.c"` and `#include "drill_render.c"` pattern (single translation unit)
- ASR client shared between GUI and headless via `shared/` directory
- All server communication is HTTP POST to localhost (configurable port)
- SSE streaming: tokens arrive via `WM_ASR_TOKEN` messages from worker thread
- Drill mode: streaming codepoints filtered to CJK-only (>= U+2E80)
