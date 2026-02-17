/*
 * voice-test-gui -- Voice transcription with graphical waveform + stability detection
 *
 * Features:
 * - Real-time graphical waveform visualization
 * - Large status counters
 * - Stability detection: sentences commit when stable
 * - Pronunciation drill mode with streaming character display
 * - TTS via Windows SAPI + server-based TTS (local-ai-server /v1/audio/speech)
 *
 * Build: clients\voice-test-gui\build.bat
 * Run:   bin\voice-test-gui.exe
 */

#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
#define _CRT_SECURE_NO_WARNINGS

#include <initguid.h>
#include <windows.h>
#include <commctrl.h>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <mferror.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <sapi.h>
#include <mmsystem.h>
#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include <ctype.h>
#include <math.h>
#include <time.h>

#include "asr_client.h"
#include "drill.h"

/* GUIDs */
DEFINE_GUID(CLSID_MMDeviceEnumerator, 0xBCDE0395, 0xE52F, 0x467C,
            0x8E, 0x3D, 0xC4, 0x57, 0x92, 0x91, 0x69, 0x2E);
DEFINE_GUID(IID_IMMDeviceEnumerator, 0xA95664D2, 0x9614, 0x4F35,
            0xA7, 0x46, 0xDE, 0x8D, 0xB6, 0x36, 0x17, 0xE6);
DEFINE_GUID(IID_IAudioEndpointVolume, 0x5CDF2C82, 0x841E, 0x4546,
            0x97, 0x22, 0x0C, 0xF7, 0x40, 0x78, 0x22, 0x9A);
DEFINE_GUID(CLSID_SpVoice, 0x96749377, 0x3391, 0x11D2,
            0x9E, 0xE3, 0x00, 0xC0, 0x4F, 0x79, 0x73, 0x96);
DEFINE_GUID(IID_ISpVoice, 0x6C44DF74, 0x72B9, 0x4992,
            0xA1, 0xEC, 0xEF, 0x99, 0x6E, 0x04, 0x22, 0xD4);

#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mf.lib")
#pragma comment(lib, "mfreadwrite.lib")
#pragma comment(lib, "mfuuid.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "winmm.lib")

#include <winhttp.h>
#include <psapi.h>
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "psapi.lib")

/* Configuration */
#define WHISPER_SAMPLE_RATE 16000
#define CHUNK_MS           2000
#define MAX_AUDIO_SECONDS  120
#define MAX_AUDIO_SAMPLES  (WHISPER_SAMPLE_RATE * MAX_AUDIO_SECONDS)
#define WAVEFORM_BARS      60
#define WAVEFORM_UPDATE_MS 50
#define STABILITY_COUNT    2
#define SILENCE_THRESHOLD  0.010f
#define SILENCE_CHUNKS     4

/* Timeline scrubbing - store energy for entire recording */
#define MAX_STORED_BARS    (MAX_AUDIO_SECONDS * 1000 / WAVEFORM_UPDATE_MS)  /* 2400 bars max */
#define SAMPLES_PER_BAR    (WHISPER_SAMPLE_RATE * WAVEFORM_UPDATE_MS / 1000) /* 800 samples */

/* Colors */
#define COLOR_BG           RGB(30, 30, 35)
#define COLOR_WAVE_LOW     RGB(50, 180, 100)
#define COLOR_WAVE_MED     RGB(180, 180, 50)
#define COLOR_WAVE_HIGH    RGB(220, 80, 80)
#define COLOR_SILENCE      RGB(80, 80, 200)
#define COLOR_TEXT         RGB(220, 220, 220)
#define COLOR_TEXT_DIM     RGB(140, 140, 140)
#define COLOR_ACCENT       RGB(100, 200, 255)
/* GUI IDs */
#define ID_BTN_RECORD     101
#define ID_TIMER_TRANSCRIBE 1
#define ID_TIMER_WAVEFORM   2
#define ID_TIMER_DEVSTATUS  3
#define ID_SCROLLBAR      106
#define ID_EDIT_CLAUDE    107
#define ID_LBL_CLAUDE     108
#define ID_EDIT_CHAT      109
#define ID_LBL_CHAT       110

/* Named pipe IPC */
#define PIPE_NAME "\\\\.\\pipe\\voice_claude"
#define WM_PIPE_RESPONSE (WM_USER + 100)
#define WM_TTS_STATUS    (WM_USER + 101)  /* wParam: 0=idle, 1=generating, 2=speaking */
#define WM_LLM_RESPONSE  (WM_USER + 102)
#define WM_TRANSCRIBE_DONE (WM_USER + 103)  /* lParam: AsrResult* */
#define WM_ASR_TOKEN       (WM_USER + 104)  /* lParam: AsrTokenMsg* */
#define WM_TTS_CACHED      (WM_USER + 105)  /* wParam: sentence idx whose word groupings are ready */
#define PIPE_BUF_SIZE 4096

/* Local LLM server */
#define LLM_SERVER_PORT   8042
#define LLM_MAX_HISTORY   20
#define LLM_MAX_CONTENT   4096
#define LLM_REQUEST_BUF   32768
#define LLM_RESPONSE_BUF  16384

/* Conversation history */
#define MAX_CHAT_LEN 16384

/* ASR server connection */
static int g_asr_port = 8090;
static const char *g_asr_language = NULL;  /* "Chinese" or NULL for auto */
static char g_asr_prompt[4096] = "";       /* context bias for continuation */

/* Resolve a path relative to the directory containing this executable.
 * Result is written to out_buf (must be MAX_PATH). Returns 0 on success. */
static int resolve_exe_relative(const char *rel_path, char *out_buf, size_t buf_size) {
    char exe_dir[MAX_PATH];
    DWORD len = GetModuleFileNameA(NULL, exe_dir, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) return -1;
    /* Strip filename to get directory */
    char *last_sep = strrchr(exe_dir, '\\');
    if (last_sep) *(last_sep + 1) = '\0';
    /* Combine and canonicalize */
    char combined[MAX_PATH * 2];
    snprintf(combined, sizeof(combined), "%s%s", exe_dir, rel_path);
    if (!GetFullPathNameA(combined, (DWORD)buf_size, out_buf, NULL)) return -1;
    return 0;
}

/* Global state */
static HWND g_hwnd_main = NULL;
static HWND g_hwnd_btn = NULL;
static HWND g_hwnd_waveform = NULL;
static HWND g_hwnd_stats = NULL;
static HWND g_hwnd_lbl_audio = NULL;
static HWND g_hwnd_scrollbar = NULL;
static HWND g_hwnd_lbl_claude = NULL;
static HWND g_hwnd_claude_response = NULL;
static HWND g_hwnd_lbl_chat = NULL;
static HWND g_hwnd_chat = NULL;
static HWND g_hwnd_sysinfo = NULL;  /* system/hardware info strip */
static HWND g_hwnd_diag = NULL;   /* full-width diagnostics strip */
static HWND g_hwnd_drill = NULL;  /* drill mode display panel */
static int g_drill_mode = 0;
static HFONT g_font_drill_chinese = NULL;
static DrillState g_drill_state;
static char g_drill_sentence_path[MAX_PATH];
static char g_drill_progress_path[MAX_PATH];
static DWORD g_drill_flash_tick = 0;  /* timestamp for diff flash effect */
#define DRILL_FLASH_MS 150
#define ID_TIMER_DRILL_FLASH 42

/* Streaming drill: accumulate codepoints + timing from SSE tokens */
static int g_drill_stream_cps[DRILL_MAX_TEXT];
static int g_drill_stream_ms[DRILL_MAX_TEXT];   /* audio_ms per codepoint */
static int g_drill_stream_len = 0;

/* Drill copy-to-clipboard visual feedback */
static DWORD g_drill_copy_tick = 0;   /* GetTickCount() of last copy */
static int   g_drill_copy_row = -1;   /* 0=target, 1=result */
#define DRILL_COPY_FLASH_MS 800
#define ID_TIMER_DRILL_COPY 43
#define ID_TIMER_PLAYBACK   44

/* TTS state */
static ISpVoice *g_tts_voice = NULL;
static int g_tts_enabled = 1;

/* Server-based TTS (local-ai-server /v1/audio/speech) */
#define TTS_SERVER_PORT 8090
#define TTS_RESPONSE_BUF (4 * 1024 * 1024)  /* 4 MB for WAV data */
static HWAVEOUT g_waveout = NULL;
static HANDLE g_waveout_done_event = NULL;
static volatile LONG g_tts_playback_ms = -1;    /* -1 = not playing, >=0 = position ms */
static volatile LONG g_tts_interrupt = 0;
static volatile HINTERNET g_tts_worker_hrequest = NULL;  /* in-flight HTTP for cancellation */
static char *g_tts_pending_text = NULL;
static int g_tts_pending_voice_idx = 0;
static int g_tts_pending_sentence_idx = -1;
static int g_tts_pending_seed = -1;  /* -1=use locked, -2=force random (tuning), >=0=explicit */
static CRITICAL_SECTION g_tts_lock;
static HANDLE g_tts_request_event = NULL;
static HANDLE g_tts_shutdown_event = NULL;
static HANDLE g_tts_thread = NULL;
static volatile int g_tts_state = 0;  /* 0=idle, 1=generating, 2=speaking, 3=error */

/* Server TTS voice presets (Qwen3-TTS) */
static const char *g_tts_voices[] = {
    "Vivian", "Serena", "Uncle_Fu", "Dylan", "Eric",
    "Ryan", "Aiden", "Ono_Anna", "Sohee"
};
#define TTS_NUM_VOICES (sizeof(g_tts_voices) / sizeof(g_tts_voices[0]))
static int g_tts_voice_idx = 0;  /* Vivian default */
static int g_tts_voice_seeds[TTS_NUM_VOICES];    /* -1 = unlocked, >=0 = locked seed */
static volatile LONG g_tts_last_seed = -1;       /* seed from last auditioned TTS */

/* Last-WAV cache: single-entry cache keyed by sentence_idx + voice_idx */
static char            *g_tts_last_wav = NULL;
static int              g_tts_last_wav_len = 0;
static int              g_tts_last_wav_sentence = -1;
static int              g_tts_last_wav_voice = -1;
static CRITICAL_SECTION g_tts_last_wav_lock;

/* TTS word timestamps (from ASR on generated audio) */
typedef struct {
    char word[64];    /* UTF-8 word text */
    int  start_ms;
    int  end_ms;
} TtsWordTimestamp;

typedef struct {
    TtsWordTimestamp *words;  /* malloc'd, NULL if none */
    int              count;
} TtsTimestamps;

/* Per-sentence word groupings (voice-independent, never purged) */
static TtsTimestamps   *g_tts_groupings       = NULL;  /* malloc'd [num_sentences] */
static int              g_tts_groupings_count  = 0;
static CRITICAL_SECTION g_tts_groupings_lock;

/* Current word timestamps (set by worker thread, read by drill renderer for playback timing) */
static TtsTimestamps    g_tts_current_ts = {0};
static CRITICAL_SECTION g_tts_current_ts_lock;

/* TTS pre-fetch thread (fetches word groupings only) */
static HANDLE           g_tts_prefetch_event     = NULL;  /* auto-reset */
static HANDLE           g_tts_prefetch_shutdown   = NULL;  /* manual-reset */
static HANDLE           g_tts_prefetch_thread     = NULL;
static volatile HINTERNET g_tts_prefetch_hrequest = NULL;  /* in-flight HTTP for cancellation */
static volatile LONG    g_tts_prefetch_priority   = -1;    /* sentence to fetch next */
static volatile LONG    g_tts_prefetch_done       = 0;     /* completed count for progress UI */
static volatile LONG    g_tts_prefetch_total      = 0;     /* total sentences for progress UI */

/* System/hardware info (queried once at startup) */
static char g_os_version[64] = "";
static char g_cpu_name[128] = "";
static char g_ram_total[32] = "";
static char g_cpu_cores[16] = "";

/* OS device status (mic/speaker volume + mute) */
static float g_mic_volume = -1.0f;   /* -1 = no device */
static int g_mic_muted = 0;
static float g_spk_volume = -1.0f;   /* -1 = no device */
static int g_spk_muted = 0;

/* Conversation history */
static char g_chat_log[MAX_CHAT_LEN] = "";
static int g_chat_len = 0;

/* Push-to-talk state */
static int g_ptt_held = 0;       /* spacebar currently held down */
static DWORD g_ptt_start_tick = 0; /* tick count when PTT started */
#define PTT_MIN_HOLD_MS 800       /* minimum hold before release stops recording */

/* Named pipe state */
static HANDLE g_pipe = INVALID_HANDLE_VALUE;
static HANDLE g_pipe_thread = NULL;
static HANDLE g_pipe_shutdown_event = NULL;  /* Signals pipe thread to exit */
static volatile int g_pipe_connected = 0;
static volatile int g_pipe_running = 0;

/* LLM mode state */
typedef enum { LLM_MODE_CLAUDE, LLM_MODE_LOCAL } LlmMode;
static volatile LlmMode g_llm_mode = LLM_MODE_CLAUDE;

/* Mandarin tutor mode */
static int g_tutor_mode = 0;
static int g_tutor_model_loaded = 0;  /* 1 = multilingual model currently loaded */

typedef struct { char role[16]; char content[LLM_MAX_CONTENT]; } LlmMessage;
static LlmMessage g_llm_history[LLM_MAX_HISTORY];
static int g_llm_history_count = 0;

/* LLM worker thread (mirrors TTS worker pattern) */
static char *g_llm_pending_prompt = NULL;
static CRITICAL_SECTION g_llm_lock;
static HANDLE g_llm_request_event = NULL;
static HANDLE g_llm_shutdown_event = NULL;
static HANDLE g_llm_thread = NULL;
static volatile LONG g_llm_interrupt = 0;
static volatile int g_llm_server_ok = 0;
static const char *g_llm_system_prompt =
    "You are a helpful voice assistant. Keep responses concise and conversational.";

static const char *g_tutor_system_prompt =
    "You are a Mandarin Chinese tutor. The student speaks in Mandarin (transcribed). "
    "Reply with this exact format for EVERY response:\n"
    "\n"
    "Chinese: (the corrected/natural Mandarin sentence using simplified characters)\n"
    "Pinyin: (full pinyin with tone marks)\n"
    "English: (English translation)\n"
    "Grammar: (one brief grammar note, if relevant)\n"
    "Prompt: (a follow-up question or prompt IN Mandarin characters to keep the conversation going)\n"
    "\n"
    "Rules:\n"
    "- Use HSK 1-2 vocabulary (beginner level)\n"
    "- Keep sentences short (under 10 words)\n"
    "- Be encouraging and patient\n"
    "- If the student's Mandarin is correct, say so briefly then give the next prompt\n"
    "- Always include all five sections";

static float g_audio_buffer[MAX_AUDIO_SAMPLES];
static int g_audio_write_pos = 0;
static int g_audio_samples = 0;
static CRITICAL_SECTION g_audio_lock;

/* Full recording buffer for WAV saving (never cleared during recording) */
static float g_recording_buffer[MAX_AUDIO_SAMPLES];
static int g_recording_samples = 0;

/* No model context needed -- transcription via HTTP to local-ai-server */

static volatile int g_capture_running = 0;
static volatile int g_capture_ready = 0;   /* 1 once first audio samples arrive */
static HANDLE g_capture_thread = NULL;
static int g_is_recording = 0;

/* Waveform data */
static float g_waveform_levels[WAVEFORM_BARS];
static float g_current_energy = 0.0f;

/* Stored waveform for entire recording (for scrubbing after stop) */
static float g_stored_levels[MAX_STORED_BARS];
static int g_stored_bar_count = 0;

/* Timeline scrubbing state */
static int g_scroll_offset = 0;          /* Which bar is at left edge of view */
static float g_marker_time = -1.0f;      /* Current marker position in seconds (-1 = none) */
static int g_marker_bar = -1;            /* Current marker position in bars */
static int g_dragging = 0;               /* Is user dragging the marker? */

/* Stability tracking */
#define MAX_HISTORY 3
static char g_transcript_history[MAX_HISTORY][8192];
static int g_history_count = 0;
static char g_finalized_text[8192] = "";
static int g_finalized_len = 0;

/* Silence detection */
static int g_silence_count = 0;
static int g_had_speech = 0;
static int g_pending_stop = 0;

/* Stats for display */
static float g_audio_seconds = 0.0f;

/* Sliding window: advances as sentences stabilize */
static int g_committed_samples = 0;  /* audio offset past stable sentences */
static int g_window_samples = 0;     /* samples in last transcription window */
static volatile int g_transcribing = 0;
static HANDLE g_transcribe_thread = NULL;  /* saved for join before context mutation */

/* Resource monitoring stats (second stats row) */
static int    g_pass_count = 0;           /* retranscription passes so far */
static double g_last_transcribe_ms = 0;   /* wall-clock time of last pass */
static double g_last_audio_window_sec = 0;/* audio duration transcribed */
static double g_last_rtf = 0;            /* realtime factor (< 1.0 = good) */
static double g_last_encode_ms = 0;      /* encoder time */
static double g_last_decode_ms = 0;      /* decoder time */
static int    g_last_common_pct = 0;     /* stability: common prefix % */
static int    g_committed_chars = 0;     /* total chars committed as [You] */
static double g_cpu_percent = 0;         /* process CPU% */
static SIZE_T g_working_set_mb = 0;      /* process working set MB */

/* CPU measurement baseline */
static ULONGLONG g_cpu_prev_kernel = 0;
static ULONGLONG g_cpu_prev_user = 0;
static ULONGLONG g_cpu_prev_time = 0;

static LARGE_INTEGER g_freq;
static FILE *g_log_file = NULL;

static HFONT g_font_large = NULL;
static HFONT g_font_medium = NULL;
static HFONT g_font_normal = NULL;
static HFONT g_font_small = NULL;
static HFONT g_font_italic = NULL;
static HBRUSH g_brush_bg = NULL;

static double get_time_ms(void) {
    LARGE_INTEGER now;
    QueryPerformanceCounter(&now);
    return (double)now.QuadPart / (double)g_freq.QuadPart * 1000.0;
}

/* Set window text from a UTF-8 string (supports CJK characters).
 * Converts UTF-8 to UTF-16 and calls SetWindowTextW. */
static void SetWindowTextUTF8(HWND hwnd, const char *utf8) {
    if (!hwnd) return;
    if (!utf8 || !utf8[0]) { SetWindowTextW(hwnd, L""); return; }
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) { SetWindowTextA(hwnd, utf8); return; }
    wchar_t *wstr = (wchar_t *)malloc(wlen * sizeof(wchar_t));
    if (!wstr) { SetWindowTextA(hwnd, utf8); return; }
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wstr, wlen);
    SetWindowTextW(hwnd, wstr);
    free(wstr);
}

/* Draw text from a UTF-8 string (supports CJK characters).
 * Converts UTF-8 to UTF-16 and calls DrawTextW. */
static int DrawTextUTF8(HDC hdc, const char *utf8, int ncount, LPRECT lprc, UINT format) {
    if (!utf8) return 0;
    int srclen = (ncount == -1) ? -1 : ncount;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, srclen, NULL, 0);
    if (wlen <= 0) return DrawTextA(hdc, utf8, ncount, lprc, format);
    wchar_t *wstr = (wchar_t *)malloc(wlen * sizeof(wchar_t));
    if (!wstr) return DrawTextA(hdc, utf8, ncount, lprc, format);
    MultiByteToWideChar(CP_UTF8, 0, utf8, srclen, wstr, wlen);
    int ret = DrawTextW(hdc, wstr, wlen, lprc, format);
    free(wstr);
    return ret;
}

static void log_event(const char *event, const char *detail) {
    if (!g_log_file) return;
    fprintf(g_log_file, "[%6.1fs] %-12s %s\n", g_audio_seconds, event, detail ? detail : "");
    fflush(g_log_file);
}


/* Forward declaration (defined later with LLM code) */
static int json_escape(const char *src, char *dst, int dst_size);

/* ---- waveOut playback (int16 PCM from server WAV) ---- */

static int g_waveout_base_sr = 0;  /* base rate before speed (always 48000 for 24kHz source) */

/* Open waveOut device at given sample rate, mono 16-bit PCM. Returns 1 on success.
 * 24kHz is upsampled to 48kHz to avoid inconsistent Windows mixer resampling. */
static int waveout_open(int sample_rate) {
    g_waveout_done_event = CreateEventA(NULL, FALSE, FALSE, NULL);
    if (!g_waveout_done_event) return 0;

    /* Open at 48kHz for 24kHz source — waveout_play_pcm does the 2x upsample */
    int base_rate = (sample_rate == 24000) ? 48000 : sample_rate;
    int device_rate = base_rate;
    g_waveout_base_sr = base_rate;

    WAVEFORMATEX wfx = {0};
    wfx.wFormatTag = WAVE_FORMAT_PCM;
    wfx.nChannels = 1;
    wfx.nSamplesPerSec = (DWORD)device_rate;
    wfx.wBitsPerSample = 16;
    wfx.nBlockAlign = wfx.nChannels * wfx.wBitsPerSample / 8;
    wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * wfx.nBlockAlign;

    MMRESULT mr = waveOutOpen(&g_waveout, WAVE_MAPPER, &wfx,
                              (DWORD_PTR)g_waveout_done_event,
                              0, CALLBACK_EVENT);
    if (mr != MMSYSERR_NOERROR) {
        CloseHandle(g_waveout_done_event);
        g_waveout_done_event = NULL;
        g_waveout = NULL;
        log_event("WAVEOUT", "waveOutOpen failed");
        return 0;
    }

    log_event("WAVEOUT", "Opened successfully");
    return 1;
}

static void waveout_close(void) {
    if (g_waveout) {
        waveOutReset(g_waveout);
        waveOutClose(g_waveout);
        g_waveout = NULL;
    }
    if (g_waveout_done_event) {
        CloseHandle(g_waveout_done_event);
        g_waveout_done_event = NULL;
    }
}

/* Play int16 PCM samples through waveOut. Returns 0 = played fully, 1 = interrupted. */
static int waveout_play_pcm(const int16_t *samples, int n_samples) {
    if (!g_waveout || n_samples <= 0) return 0;

    /* Upsample to 48kHz if source is 24kHz to avoid inconsistent Windows
     * audio mixer resampling (24kHz is non-standard, 48kHz is native). */
    int16_t *upsampled = NULL;
    if (g_waveout_base_sr == 48000) {
        upsampled = (int16_t *)malloc((size_t)n_samples * 2 * sizeof(int16_t));
        if (upsampled) {
            for (int i = 0; i < n_samples; i++) {
                upsampled[i * 2]     = samples[i];
                upsampled[i * 2 + 1] = samples[i];
            }
            samples = upsampled;
            n_samples *= 2;
        }
    }

    /* Submit entire buffer as one WAVEHDR */
    int interrupted = 0;
    WAVEHDR hdr = {0};
    hdr.lpData = (LPSTR)samples;
    hdr.dwBufferLength = (DWORD)(n_samples * sizeof(int16_t));

    MMRESULT mr = waveOutPrepareHeader(g_waveout, &hdr, sizeof(hdr));
    if (mr != MMSYSERR_NOERROR) return 0;

    ResetEvent(g_waveout_done_event);
    InterlockedExchange(&g_tts_playback_ms, 0);
    mr = waveOutWrite(g_waveout, &hdr, sizeof(hdr));
    if (mr != MMSYSERR_NOERROR) {
        InterlockedExchange(&g_tts_playback_ms, -1);
        waveOutUnprepareHeader(g_waveout, &hdr, sizeof(hdr));
        return 0;
    }

    /* Poll for completion or interrupt */
    while (!(hdr.dwFlags & WHDR_DONE)) {
        DWORD ret = WaitForSingleObject(g_waveout_done_event, 50);
        if (ret == WAIT_TIMEOUT || ret == WAIT_OBJECT_0) {
            if (InterlockedCompareExchange(&g_tts_interrupt, 0, 0)) {
                waveOutReset(g_waveout);
                interrupted = 1;
                break;
            }
            /* Update playback position for karaoke highlight.
             * Use base rate (48000) not device rate (48000*speed) so
             * position maps to original audio ms regardless of speed. */
            if (g_waveout_base_sr > 0) {
                MMTIME mmt = {0};
                mmt.wType = TIME_SAMPLES;
                if (waveOutGetPosition(g_waveout, &mmt, sizeof(mmt)) == MMSYSERR_NOERROR
                    && mmt.wType == TIME_SAMPLES) {
                    InterlockedExchange(&g_tts_playback_ms,
                        (LONG)((double)mmt.u.sample * 1000.0 / g_waveout_base_sr));
                }
            }
        }
    }

    InterlockedExchange(&g_tts_playback_ms, -1);
    waveOutUnprepareHeader(g_waveout, &hdr, sizeof(hdr));
    free(upsampled);
    return interrupted;
}

/* ---- Base64 decoder (for TTS timestamp responses) ---- */

static const int b64_decode_table[256] = {
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,62,-1,-1,-1,63,
    52,53,54,55,56,57,58,59,60,61,-1,-1,-1,-1,-1,-1,
    -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14,
    15,16,17,18,19,20,21,22,23,24,25,-1,-1,-1,-1,-1,
    -1,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,
    41,42,43,44,45,46,47,48,49,50,51,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
    -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,
};

static char *base64_decode(const char *b64, int b64_len, int *out_len) {
    *out_len = 0;
    int pad = 0;
    if (b64_len >= 1 && b64[b64_len - 1] == '=') pad++;
    if (b64_len >= 2 && b64[b64_len - 2] == '=') pad++;
    int out_size = (b64_len / 4) * 3 - pad;
    if (out_size <= 0) return NULL;
    char *out = (char *)malloc(out_size);
    if (!out) return NULL;

    int j = 0;
    for (int i = 0; i + 3 < b64_len; i += 4) {
        int a = b64_decode_table[(unsigned char)b64[i]];
        int b = b64_decode_table[(unsigned char)b64[i+1]];
        int c = b64_decode_table[(unsigned char)b64[i+2]];
        int d = b64_decode_table[(unsigned char)b64[i+3]];
        if (a < 0 || b < 0) { free(out); return NULL; }
        int v = (a << 18) | (b << 12);
        if (c >= 0) v |= (c << 6);
        if (d >= 0) v |= d;
        if (j < out_size) out[j++] = (char)((v >> 16) & 0xFF);
        if (j < out_size && c >= 0) out[j++] = (char)((v >> 8) & 0xFF);
        if (j < out_size && d >= 0) out[j++] = (char)(v & 0xFF);
    }
    *out_len = j;
    return out;
}

/* ---- TTS timestamp JSON parser ---- */

/* Safe strnstr for non-null-terminated buffers */
static const char *strnstr_safe(const char *haystack, const char *needle, int haystack_len) {
    int needle_len = (int)strlen(needle);
    for (int i = 0; i <= haystack_len - needle_len; i++) {
        if (memcmp(haystack + i, needle, needle_len) == 0)
            return haystack + i;
    }
    return NULL;
}

/* Find a JSON string field value. Returns pointer to the opening quote content,
 * sets *value_len to content length (between quotes). Returns NULL if not found. */
static const char *json_find_string(const char *json, int json_len,
                                     const char *key, int *value_len) {
    char pattern[128];
    snprintf(pattern, sizeof(pattern), "\"%s\":", key);
    const char *p = strnstr_safe(json, pattern, json_len);
    if (!p) return NULL;
    p += strlen(pattern);
    while (p < json + json_len && (*p == ' ' || *p == '\t')) p++;
    if (p >= json + json_len || *p != '"') return NULL;
    p++; /* skip opening quote */
    const char *start = p;
    while (p < json + json_len && *p != '"') {
        if (*p == '\\') p++; /* skip escaped char */
        p++;
    }
    *value_len = (int)(p - start);
    return start;
}

static double json_find_double(const char *json, int json_len,
                                const char *key, double fallback) {
    char pattern[128];
    snprintf(pattern, sizeof(pattern), "\"%s\":", key);
    const char *p = strnstr_safe(json, pattern, json_len);
    if (!p) return fallback;
    p += strlen(pattern);
    while (p < json + json_len && (*p == ' ' || *p == '\t')) p++;
    int remaining = (int)(json + json_len - p);
    if (remaining <= 0) return fallback;
    char tmp[64];
    if (remaining > 63) remaining = 63;
    memcpy(tmp, p, remaining);
    tmp[remaining] = '\0';
    return atof(tmp);
}

/* Parse TTS timestamp JSON response.
 * Returns 0 on success, -1 on error. Caller must free *wav_out. */
static int tts_parse_timestamp_response(const char *json, int json_len,
                                         char **wav_out, int *wav_len_out,
                                         TtsTimestamps *ts_out) {
    *wav_out = NULL;
    *wav_len_out = 0;
    ts_out->words = NULL;
    ts_out->count = 0;

    /* Extract "audio":"<base64>" */
    int audio_b64_len = 0;
    const char *audio_b64 = json_find_string(json, json_len, "audio", &audio_b64_len);
    if (!audio_b64 || audio_b64_len <= 0) return -1;

    int wav_len = 0;
    char *wav = base64_decode(audio_b64, audio_b64_len, &wav_len);
    if (!wav || wav_len <= 0) return -1;

    *wav_out = wav;
    *wav_len_out = wav_len;

    /* Parse "words":[...] array */
    const char *words_key = strnstr_safe(json, "\"words\":", json_len);
    if (!words_key) return 0; /* no words array is OK */

    const char *arr_start = strchr(words_key, '[');
    if (!arr_start) return 0;
    arr_start++;

    /* Count objects to allocate */
    int count = 0;
    {
        const char *scan = arr_start;
        const char *end = json + json_len;
        while (scan < end && *scan != ']') {
            if (*scan == '{') count++;
            scan++;
        }
    }
    if (count <= 0) return 0;

    TtsWordTimestamp *words = (TtsWordTimestamp *)calloc(count, sizeof(TtsWordTimestamp));
    if (!words) return 0;

    /* Parse each word object */
    const char *p = arr_start;
    const char *end = json + json_len;
    for (int i = 0; i < count && p < end; i++) {
        /* Find next '{' */
        while (p < end && *p != '{') p++;
        if (p >= end) break;

        /* Find matching '}' */
        const char *obj_start = p;
        const char *obj_end = strchr(p, '}');
        if (!obj_end) break;
        int obj_len = (int)(obj_end - obj_start + 1);

        /* Extract "word" */
        int wlen = 0;
        const char *wval = json_find_string(obj_start, obj_len, "word", &wlen);
        if (wval && wlen > 0) {
            int copy_len = wlen < 63 ? wlen : 63;
            memcpy(words[i].word, wval, copy_len);
            words[i].word[copy_len] = '\0';
        }

        /* Extract "start" and "end" as seconds, convert to ms */
        double start_s = json_find_double(obj_start, obj_len, "start", 0.0);
        double end_s = json_find_double(obj_start, obj_len, "end", 0.0);
        words[i].start_ms = (int)(start_s * 1000.0);
        words[i].end_ms = (int)(end_s * 1000.0);

        p = obj_end + 1;
    }

    ts_out->words = words;
    ts_out->count = count;
    return 0;
}

/* ---- TTS HTTP client (local-ai-server /v1/audio/speech) ---- */

/* POST to TTS server, receive WAV bytes. Returns 0 on success, -1 on failure.
 * Caller must free *wav_out.
 * ts_out (optional): when non-NULL, sends "timestamps":true, parses JSON response
 * to extract WAV + word timestamps. When NULL, receives raw WAV bytes.
 * cancel_handle (optional): published for external cancellation via
 * WinHttpCloseHandle from another thread. InterlockedExchangePointer ensures
 * exactly one side (caller or canceller) closes the request handle. */
static int tts_request(const char *text, const char *voice, int seed,
                       char **wav_out, int *wav_len,
                       TtsTimestamps *ts_out, int *seed_out,
                       volatile HINTERNET *cancel_handle) {
    *wav_out = NULL;
    *wav_len = 0;
    if (ts_out) { ts_out->words = NULL; ts_out->count = 0; }
    if (seed_out) *seed_out = -1;

    HINTERNET hSession = WinHttpOpen(L"VoiceNoteGUI/1.0",
                                      WINHTTP_ACCESS_TYPE_NO_PROXY,
                                      WINHTTP_NO_PROXY_NAME,
                                      WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return -1;

    HINTERNET hConnect = WinHttpConnect(hSession, L"localhost",
                                         TTS_SERVER_PORT, 0);
    if (!hConnect) {
        WinHttpCloseHandle(hSession);
        return -1;
    }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
                                             L"/v1/audio/speech",
                                             NULL, WINHTTP_NO_REFERER,
                                             WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return -1;
    }

    /* Publish handle so another thread can close it to abort blocking I/O */
    if (cancel_handle)
        InterlockedExchangePointer((volatile PVOID *)cancel_handle, hRequest);

    /* 5s connect, 60s receive (TTS generation can be slow) */
    WinHttpSetTimeouts(hRequest, 5000, 5000, 60000, 60000);

    /* Build JSON body */
    char escaped[4096];
    json_escape(text, escaped, sizeof(escaped));
    char body[8192];
    if (ts_out) {
        if (seed >= 0) {
            snprintf(body, sizeof(body),
                     "{\"input\":\"%s\",\"voice\":\"%s\",\"response_format\":\"wav\","
                     "\"timestamps\":true,\"language\":\"Chinese\",\"seed\":%d}",
                     escaped, voice, seed);
        } else {
            snprintf(body, sizeof(body),
                     "{\"input\":\"%s\",\"voice\":\"%s\",\"response_format\":\"wav\","
                     "\"timestamps\":true,\"language\":\"Chinese\"}",
                     escaped, voice);
        }
    } else {
        if (seed >= 0) {
            snprintf(body, sizeof(body),
                     "{\"input\":\"%s\",\"voice\":\"%s\",\"response_format\":\"wav\",\"seed\":%d}",
                     escaped, voice, seed);
        } else {
            snprintf(body, sizeof(body),
                     "{\"input\":\"%s\",\"voice\":\"%s\",\"response_format\":\"wav\"}",
                     escaped, voice);
        }
    }

    DWORD body_len = (DWORD)strlen(body);
    BOOL ok = WinHttpSendRequest(hRequest,
                                  L"Content-Type: application/json\r\n",
                                  (DWORD)-1L,
                                  body, body_len, body_len, 0);

    if (ok) ok = WinHttpReceiveResponse(hRequest, NULL);

    int result = -1;
    if (ok) {
        /* Check HTTP status */
        DWORD status_code = 0;
        DWORD sz = sizeof(status_code);
        WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                            NULL, &status_code, &sz, NULL);
        if (status_code != 200) {
            char errmsg[64];
            snprintf(errmsg, sizeof(errmsg), "TTS server returned HTTP %lu", status_code);
            log_event("TTS_SRV", errmsg);
        } else {
            /* Read response data */
            char *buf = (char *)malloc(TTS_RESPONSE_BUF);
            if (buf) {
                DWORD total = 0, bytes_read = 0;
                while (WinHttpReadData(hRequest, buf + total,
                                        TTS_RESPONSE_BUF - total, &bytes_read)) {
                    if (bytes_read == 0) break;
                    total += bytes_read;
                    if (total >= (DWORD)TTS_RESPONSE_BUF) break;
                }
                if (total > 0) {
                    if (ts_out) {
                        /* JSON response: parse base64 audio + word timestamps */
                        if (seed_out) {
                            *seed_out = (int)json_find_double(buf, (int)total, "seed", -1.0);
                        }
                        int rc = tts_parse_timestamp_response(buf, (int)total,
                                                               wav_out, wav_len, ts_out);
                        free(buf);
                        if (rc == 0 && *wav_out) result = 0;
                    } else {
                        /* Raw WAV response */
                        *wav_out = buf;
                        *wav_len = (int)total;
                        result = 0;
                    }
                } else {
                    free(buf);
                }
            }
        }
    }

    /* Reclaim handle: if external cancel already closed it, skip our close.
     * Both sides race to swap the slot to NULL; the winner closes. */
    if (cancel_handle) {
        if (!InterlockedExchangePointer((volatile PVOID *)cancel_handle, NULL))
            hRequest = NULL;
    }
    if (hRequest) WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);
    return result;
}

/* Parse WAV header to find PCM data. Returns 0 on success.
 * Sets *pcm_out to point into wav_data (not a new allocation),
 * *n_samples to sample count, *sample_rate to Hz. */
static int wav_parse_header(const char *wav_data, int wav_len,
                            const int16_t **pcm_out, int *n_samples, int *sample_rate) {
    if (wav_len < 44) return -1;
    if (memcmp(wav_data, "RIFF", 4) != 0) return -1;
    if (memcmp(wav_data + 8, "WAVE", 4) != 0) return -1;

    /* Walk chunks to find "fmt " and "data" */
    int fmt_found = 0;
    int sr = 0, bits = 0, channels = 0;
    int pos = 12;
    while (pos + 8 <= wav_len) {
        const char *chunk_id = wav_data + pos;
        int chunk_size = *(const int32_t *)(wav_data + pos + 4);
        if (chunk_size < 0 || pos + 8 + chunk_size > wav_len) break;

        if (memcmp(chunk_id, "fmt ", 4) == 0 && chunk_size >= 16) {
            int16_t format = *(const int16_t *)(wav_data + pos + 8);
            if (format != 1) return -1;  /* not PCM */
            channels = *(const int16_t *)(wav_data + pos + 10);
            sr = *(const int32_t *)(wav_data + pos + 12);
            bits = *(const int16_t *)(wav_data + pos + 22);
            fmt_found = 1;
        } else if (memcmp(chunk_id, "data", 4) == 0 && fmt_found) {
            if (bits != 16 || channels != 1) return -1;  /* only mono 16-bit */
            *pcm_out = (const int16_t *)(wav_data + pos + 8);
            *n_samples = chunk_size / sizeof(int16_t);
            *sample_rate = sr;
            return 0;
        }

        pos += 8 + chunk_size;
        if (chunk_size & 1) pos++;  /* chunks are 2-byte aligned */
    }
    return -1;
}

/* ---- Per-sentence word grouping cache (voice-independent, permanent) ---- */

static void tts_groupings_init(int n) {
    InitializeCriticalSection(&g_tts_groupings_lock);
    g_tts_groupings = (TtsTimestamps *)calloc(n, sizeof(TtsTimestamps));
    g_tts_groupings_count = n;
}

static void tts_groupings_destroy(void) {
    if (!g_tts_groupings) return;
    for (int i = 0; i < g_tts_groupings_count; i++)
        free(g_tts_groupings[i].words);
    free(g_tts_groupings);
    g_tts_groupings = NULL;
    g_tts_groupings_count = 0;
    DeleteCriticalSection(&g_tts_groupings_lock);
}

static int tts_groupings_has(int idx) {
    if (idx < 0 || idx >= g_tts_groupings_count) return 0;
    int found;
    EnterCriticalSection(&g_tts_groupings_lock);
    found = (g_tts_groupings[idx].count > 0) ? 1 : 0;
    LeaveCriticalSection(&g_tts_groupings_lock);
    return found;
}

/* Copy grouping timestamps for sentence idx. Returns 1 on hit. Caller frees ts->words. */
static int tts_groupings_copy(int idx, TtsTimestamps *ts) {
    if (idx < 0 || idx >= g_tts_groupings_count || !ts) return 0;
    int found = 0;
    EnterCriticalSection(&g_tts_groupings_lock);
    TtsTimestamps *e = &g_tts_groupings[idx];
    if (e->count > 0 && e->words) {
        ts->count = e->count;
        ts->words = (TtsWordTimestamp *)malloc(e->count * sizeof(TtsWordTimestamp));
        if (ts->words) {
            memcpy(ts->words, e->words, e->count * sizeof(TtsWordTimestamp));
            found = 1;
        } else {
            ts->count = 0;
        }
    } else {
        ts->words = NULL;
        ts->count = 0;
    }
    LeaveCriticalSection(&g_tts_groupings_lock);
    return found;
}

/* Store grouping (copies timestamps). */
static void tts_groupings_put(int idx, const TtsTimestamps *ts) {
    if (idx < 0 || idx >= g_tts_groupings_count) return;
    if (!ts || ts->count <= 0 || !ts->words) return;
    EnterCriticalSection(&g_tts_groupings_lock);
    TtsTimestamps *e = &g_tts_groupings[idx];
    free(e->words);
    e->count = ts->count;
    e->words = (TtsWordTimestamp *)malloc(ts->count * sizeof(TtsWordTimestamp));
    if (e->words)
        memcpy(e->words, ts->words, ts->count * sizeof(TtsWordTimestamp));
    else
        e->count = 0;
    LeaveCriticalSection(&g_tts_groupings_lock);
}

/* ---- Grouping disk persistence (voice-independent, permanent) ---- */

static uint32_t fnv1a(const char *s) {
    uint32_t h = 2166136261u;
    for (; *s; s++) { h ^= (unsigned char)*s; h *= 16777619u; }
    return h;
}

/* Build grouping disk path: <exe_dir>\..\tts_cache\groupings\<hash>.ts
 * Creates directories if missing. Returns 0 on success. */
static int tts_grouping_disk_path(const char *chinese, char *out, int out_size) {
    char base[MAX_PATH];
    if (resolve_exe_relative("..\\tts_cache", base, sizeof(base)) != 0) return -1;
    CreateDirectoryA(base, NULL);
    char grp_dir[MAX_PATH];
    snprintf(grp_dir, sizeof(grp_dir), "%s\\groupings", base);
    CreateDirectoryA(grp_dir, NULL);
    uint32_t hash = fnv1a(chinese);
    snprintf(out, out_size, "%s\\%08x.ts", grp_dir, hash);
    return 0;
}

static void tts_grouping_disk_save(const char *chinese, const TtsTimestamps *ts) {
    if (!ts || ts->count <= 0 || !ts->words) return;
    char path[MAX_PATH];
    if (tts_grouping_disk_path(chinese, path, sizeof(path)) != 0) return;
    FILE *f = fopen(path, "w");
    if (!f) return;
    for (int i = 0; i < ts->count; i++)
        fprintf(f, "%s\t%d\t%d\n", ts->words[i].word,
                ts->words[i].start_ms, ts->words[i].end_ms);
    fclose(f);
}

/* Returns 1 on hit. Caller owns ts->words. */
static int tts_grouping_disk_load(const char *chinese, TtsTimestamps *ts) {
    ts->words = NULL;
    ts->count = 0;
    char path[MAX_PATH];
    if (tts_grouping_disk_path(chinese, path, sizeof(path)) != 0) return 0;
    FILE *f = fopen(path, "r");
    if (!f) return 0;
    int nlines = 0;
    char line[256];
    while (fgets(line, sizeof(line), f)) nlines++;
    if (nlines <= 0) { fclose(f); return 0; }
    rewind(f);
    ts->words = (TtsWordTimestamp *)malloc(nlines * sizeof(TtsWordTimestamp));
    if (!ts->words) { fclose(f); return 0; }
    int n = 0;
    while (fgets(line, sizeof(line), f) && n < nlines) {
        char word[64] = "";
        int s = 0, e = 0;
        if (sscanf(line, "%63[^\t]\t%d\t%d", word, &s, &e) == 3) {
            strncpy(ts->words[n].word, word, sizeof(ts->words[n].word) - 1);
            ts->words[n].word[sizeof(ts->words[n].word) - 1] = '\0';
            ts->words[n].start_ms = s;
            ts->words[n].end_ms = e;
            n++;
        }
    }
    ts->count = n;
    fclose(f);
    if (n == 0) { free(ts->words); ts->words = NULL; return 0; }
    return 1;
}

/* ---- Per-voice seed persistence (tts_cache/seeds.txt) ---- */

static void tts_seeds_save(void) {
    char path[MAX_PATH];
    if (resolve_exe_relative("..\\tts_cache", path, sizeof(path)) != 0) return;
    CreateDirectoryA(path, NULL);
    char fpath[MAX_PATH];
    snprintf(fpath, sizeof(fpath), "%s\\seeds.txt", path);
    FILE *f = fopen(fpath, "w");
    if (!f) return;
    for (int i = 0; i < (int)TTS_NUM_VOICES; i++) {
        if (g_tts_voice_seeds[i] >= 0) {
            fprintf(f, "%s\t%d\n", g_tts_voices[i], g_tts_voice_seeds[i]);
        }
    }
    fclose(f);
}

static void tts_seeds_load(void) {
    /* Init all to -1 (unlocked) */
    for (int i = 0; i < (int)TTS_NUM_VOICES; i++)
        g_tts_voice_seeds[i] = -1;

    char path[MAX_PATH];
    if (resolve_exe_relative("..\\tts_cache\\seeds.txt", path, sizeof(path)) != 0) return;
    FILE *f = fopen(path, "r");
    if (!f) return;
    char line[128];
    while (fgets(line, sizeof(line), f)) {
        char name[64];
        int seed_val = -1;
        if (sscanf(line, "%63[^\t]\t%d", name, &seed_val) == 2) {
            for (int i = 0; i < (int)TTS_NUM_VOICES; i++) {
                if (_stricmp(g_tts_voices[i], name) == 0) {
                    g_tts_voice_seeds[i] = seed_val;
                    break;
                }
            }
        }
    }
    fclose(f);
}

/* ---- TTS worker thread (server-based) ---- */

static void tts_last_wav_clear(void) {
    EnterCriticalSection(&g_tts_last_wav_lock);
    free(g_tts_last_wav);
    g_tts_last_wav = NULL;
    g_tts_last_wav_len = 0;
    g_tts_last_wav_sentence = -1;
    g_tts_last_wav_voice = -1;
    LeaveCriticalSection(&g_tts_last_wav_lock);
}

static DWORD WINAPI tts_worker_proc(LPVOID param) {
    (void)param;
    HANDLE events[2] = { g_tts_request_event, g_tts_shutdown_event };

    while (1) {
        DWORD which = WaitForMultipleObjects(2, events, FALSE, INFINITE);
        if (which != WAIT_OBJECT_0) break; /* shutdown */

        /* Grab pending text, voice, sentence index, and seed */
        EnterCriticalSection(&g_tts_lock);
        char *text = g_tts_pending_text;
        g_tts_pending_text = NULL;
        int voice_idx = g_tts_pending_voice_idx;
        int sentence_idx = g_tts_pending_sentence_idx;
        int pending_seed = g_tts_pending_seed;
        LeaveCriticalSection(&g_tts_lock);

        if (!text) continue;

        /* Clear interrupt flag before request */
        InterlockedExchange(&g_tts_interrupt, 0);

        /* Compute effective seed:
         * pending_seed == -2: tuning mode (force random)
         * pending_seed == -1: use voice's locked seed, or random if unlocked
         * pending_seed >= 0:  use this specific seed */
        int effective_seed;
        if (pending_seed == -2) {
            effective_seed = -1;  /* random */
        } else if (pending_seed == -1) {
            effective_seed = g_tts_voice_seeds[voice_idx];  /* -1 if unlocked */
        } else {
            effective_seed = pending_seed;
        }

        int want_ts = g_drill_mode;
        TtsTimestamps worker_ts = {0};
        char *wav_data = NULL;
        int wav_len = 0;
        int from_cache = 0;

        /* Check last-WAV cache (skip for tuning: pending_seed == -2) */
        if (pending_seed != -2 && sentence_idx >= 0) {
            EnterCriticalSection(&g_tts_last_wav_lock);
            if (g_tts_last_wav && g_tts_last_wav_len > 0
                && g_tts_last_wav_sentence == sentence_idx
                && g_tts_last_wav_voice == voice_idx) {
                wav_data = (char *)malloc(g_tts_last_wav_len);
                if (wav_data) {
                    memcpy(wav_data, g_tts_last_wav, g_tts_last_wav_len);
                    wav_len = g_tts_last_wav_len;
                    from_cache = 1;
                }
            }
            LeaveCriticalSection(&g_tts_last_wav_lock);
            if (from_cache) {
                log_event("TTS_SRV", "Replay from cache");
                free(text);
                tts_groupings_copy(sentence_idx, &worker_ts);
            }
        }

        /* Fetch from server if not cached */
        if (!from_cache) {
            log_event("TTS_SRV", "Requesting speech...");
            PostMessageA(g_hwnd_main, WM_TTS_STATUS, 1, 0); /* generating */

            int seed_out = -1;
            int rc = tts_request(text, g_tts_voices[voice_idx], effective_seed,
                                 &wav_data, &wav_len,
                                 want_ts ? &worker_ts : NULL,
                                 want_ts ? &seed_out : NULL,
                                 &g_tts_worker_hrequest);

            if (rc != 0 || !wav_data) {
                log_event("TTS_SRV", "Request failed");
                PostMessageA(g_hwnd_main, WM_TTS_STATUS, 3, 0); /* error */
                free(text);
                free(wav_data);
                free(worker_ts.words);
                continue;
            }

            /* Store returned seed for UI display */
            if (seed_out >= 0) {
                InterlockedExchange(&g_tts_last_seed, (LONG)seed_out);
            }

            /* Check interrupt after network request */
            if (InterlockedCompareExchange(&g_tts_interrupt, 0, 0)) {
                log_event("TTS_SRV", "Interrupted after download");
                free(text);
                free(wav_data);
                free(worker_ts.words);
                PostMessageA(g_hwnd_main, WM_TTS_STATUS, 0, 0);
                continue;
            }

            /* Store word groupings (voice-independent, permanent) */
            if (sentence_idx >= 0 && want_ts && worker_ts.count > 0) {
                tts_groupings_put(sentence_idx, &worker_ts);
                tts_grouping_disk_save(text, &worker_ts);
            }
            free(text);

            /* Store in last-WAV cache */
            if (sentence_idx >= 0 && wav_data && wav_len > 0) {
                EnterCriticalSection(&g_tts_last_wav_lock);
                free(g_tts_last_wav);
                g_tts_last_wav = (char *)malloc(wav_len);
                if (g_tts_last_wav) {
                    memcpy(g_tts_last_wav, wav_data, wav_len);
                    g_tts_last_wav_len = wav_len;
                    g_tts_last_wav_sentence = sentence_idx;
                    g_tts_last_wav_voice = voice_idx;
                } else {
                    g_tts_last_wav_len = 0;
                    g_tts_last_wav_sentence = -1;
                    g_tts_last_wav_voice = -1;
                }
                LeaveCriticalSection(&g_tts_last_wav_lock);
            }
        }

        /* Publish word timestamps for drill renderer */
        EnterCriticalSection(&g_tts_current_ts_lock);
        free(g_tts_current_ts.words);
        g_tts_current_ts = worker_ts;
        LeaveCriticalSection(&g_tts_current_ts_lock);

        /* Parse WAV header */
        const int16_t *pcm = NULL;
        int n_samples = 0, sr = 0;
        if (wav_parse_header(wav_data, wav_len, &pcm, &n_samples, &sr) != 0) {
            log_event("TTS_SRV", "Failed to parse WAV header");
            free(wav_data);
            PostMessageA(g_hwnd_main, WM_TTS_STATUS, 0, 0);
            continue;
        }

        char msg[80];
        snprintf(msg, sizeof(msg), "%s %d samples at %d Hz (%.1fs)",
                 from_cache ? "Cached" : "Received",
                 n_samples, sr, (double)n_samples / sr);
        log_event("TTS_SRV", msg);

        /* Open waveOut at the WAV sample rate, play, close */
        if (waveout_open(sr)) {
            PostMessageA(g_hwnd_main, WM_TTS_STATUS, 2, 0); /* speaking */
            int was_interrupted = waveout_play_pcm(pcm, n_samples);
            waveout_close();
            if (was_interrupted) {
                log_event("TTS_SRV", "Playback interrupted");
            }
        } else {
            log_event("TTS_SRV", "waveout_open failed");
        }

        free(wav_data);
        PostMessageA(g_hwnd_main, WM_TTS_STATUS, 0, 0); /* idle */
    }

    return 0;
}

static int tts_worker_start(void) {
    InitializeCriticalSection(&g_tts_lock);
    InitializeCriticalSection(&g_tts_current_ts_lock);
    InitializeCriticalSection(&g_tts_last_wav_lock);
    g_tts_request_event = CreateEventA(NULL, FALSE, FALSE, NULL);   /* auto-reset */
    g_tts_shutdown_event = CreateEventA(NULL, TRUE, FALSE, NULL);   /* manual-reset */
    if (!g_tts_request_event || !g_tts_shutdown_event) {
        log_event("TTS_SRV", "Failed to create worker events");
        return 0;
    }

    g_tts_thread = CreateThread(NULL, 0, tts_worker_proc, NULL, 0, NULL);
    if (!g_tts_thread) {
        log_event("TTS_SRV", "Failed to create worker thread");
        return 0;
    }

    log_event("TTS_SRV", "Worker thread started");
    return 1;
}

static void tts_worker_stop(void) {
    /* Signal interrupt + shutdown */
    InterlockedExchange(&g_tts_interrupt, 1);
    /* Cancel any in-flight HTTP request so thread unblocks immediately */
    {
        HINTERNET h = InterlockedExchangePointer(
            (volatile PVOID *)&g_tts_worker_hrequest, NULL);
        if (h) WinHttpCloseHandle(h);
    }
    if (g_tts_shutdown_event) {
        SetEvent(g_tts_shutdown_event);
    }
    if (g_tts_thread) {
        WaitForSingleObject(g_tts_thread, 5000);
        CloseHandle(g_tts_thread);
        g_tts_thread = NULL;
    }
    waveout_close();
    if (g_tts_request_event) {
        CloseHandle(g_tts_request_event);
        g_tts_request_event = NULL;
    }
    if (g_tts_shutdown_event) {
        CloseHandle(g_tts_shutdown_event);
        g_tts_shutdown_event = NULL;
    }
    EnterCriticalSection(&g_tts_lock);
    free(g_tts_pending_text);
    g_tts_pending_text = NULL;
    LeaveCriticalSection(&g_tts_lock);
    DeleteCriticalSection(&g_tts_lock);
    EnterCriticalSection(&g_tts_current_ts_lock);
    free(g_tts_current_ts.words);
    g_tts_current_ts.words = NULL;
    g_tts_current_ts.count = 0;
    LeaveCriticalSection(&g_tts_current_ts_lock);
    DeleteCriticalSection(&g_tts_current_ts_lock);
    EnterCriticalSection(&g_tts_last_wav_lock);
    free(g_tts_last_wav);
    g_tts_last_wav = NULL;
    g_tts_last_wav_len = 0;
    LeaveCriticalSection(&g_tts_last_wav_lock);
    DeleteCriticalSection(&g_tts_last_wav_lock);
}

/* ---- TTS pre-fetch thread (word groupings only, voice-independent) ---- */

/* Fetch word groupings for one sentence. Checks disk cache first, then server.
 * Discards audio — only keeps timestamps. Returns 1 on success. */
static int tts_prefetch_fetch_one(int idx) {
    if (idx < 0 || idx >= g_drill_state.num_sentences) return 0;
    const char *text = g_drill_state.sentences[idx].chinese;
    if (!text[0]) return 0;

    TtsTimestamps ts = {0};

    /* Try disk cache first */
    if (tts_grouping_disk_load(text, &ts)) {
        tts_groupings_put(idx, &ts);
        free(ts.words);
        InterlockedIncrement(&g_tts_prefetch_done);
        char pmsg[80];
        snprintf(pmsg, sizeof(pmsg), "Grouping from disk: sentence %d (%ld/%ld)",
                 idx,
                 InterlockedCompareExchange(&g_tts_prefetch_done, 0, 0),
                 InterlockedCompareExchange(&g_tts_prefetch_total, 0, 0));
        log_event("TTS_PRE", pmsg);
        PostMessageA(g_hwnd_main, WM_TTS_CACHED, (WPARAM)idx, 0);
        return 1;
    }

    /* Fetch from server — use first voice, no seed (groupings are voice-independent) */
    char *wav_data = NULL;
    int wav_len = 0;
    int rc = tts_request(text, g_tts_voices[0], -1,
                         &wav_data, &wav_len,
                         &ts, NULL, &g_tts_prefetch_hrequest);
    free(wav_data);  /* discard audio — we only want groupings */
    if (rc != 0 || ts.count <= 0) {
        free(ts.words);
        return 0;
    }

    tts_grouping_disk_save(text, &ts);
    tts_groupings_put(idx, &ts);
    free(ts.words);
    InterlockedIncrement(&g_tts_prefetch_done);

    char pmsg[80];
    snprintf(pmsg, sizeof(pmsg), "Prefetched grouping: sentence %d (%ld/%ld)",
             idx,
             InterlockedCompareExchange(&g_tts_prefetch_done, 0, 0),
             InterlockedCompareExchange(&g_tts_prefetch_total, 0, 0));
    log_event("TTS_PRE", pmsg);
    PostMessageA(g_hwnd_main, WM_TTS_CACHED, (WPARAM)idx, 0);
    return 1;
}

static DWORD WINAPI tts_prefetch_proc(LPVOID param) {
    (void)param;
    HANDLE events[2] = { g_tts_prefetch_event, g_tts_prefetch_shutdown };

    while (1) {
        DWORD which = WaitForMultipleObjects(2, events, FALSE, INFINITE);
        if (which != WAIT_OBJECT_0) break; /* shutdown */

        int n = g_tts_groupings_count;

        /* Count already-cached entries to init done counter */
        LONG already = 0;
        for (int i = 0; i < n; i++) {
            if (tts_groupings_has(i)) already++;
        }
        InterlockedExchange(&g_tts_prefetch_done, already);
        InterlockedExchange(&g_tts_prefetch_total, (LONG)n);

        /* Sequential sweep 0..N with priority interrupts */
        for (int i = 0; i < n; i++) {
            if (WaitForSingleObject(g_tts_prefetch_shutdown, 0) == WAIT_OBJECT_0)
                goto done;

            /* Check priority: service it first, then resume */
            LONG pri = InterlockedExchange(&g_tts_prefetch_priority, -1);
            if (pri >= 0 && pri < n && !tts_groupings_has((int)pri)) {
                if (!tts_prefetch_fetch_one((int)pri)) {
                    if (WaitForSingleObject(g_tts_prefetch_shutdown, 2000) == WAIT_OBJECT_0)
                        goto done;
                }
            }

            /* Skip already-cached */
            if (tts_groupings_has(i)) continue;

            /* Fetch this sentence's groupings */
            if (!tts_prefetch_fetch_one(i)) {
                if (WaitForSingleObject(g_tts_prefetch_shutdown, 2000) == WAIT_OBJECT_0)
                    goto done;
                i--;  /* retry */
                continue;
            }
        }
        /* Sweep complete — wait for next event */
    }

done:
    return 0;
}

static void tts_prefetch_start(void) {
    InterlockedExchange(&g_tts_prefetch_priority, -1);

    g_tts_prefetch_event = CreateEventA(NULL, FALSE, FALSE, NULL);     /* auto-reset */
    g_tts_prefetch_shutdown = CreateEventA(NULL, TRUE, FALSE, NULL);   /* manual-reset */
    if (!g_tts_prefetch_event || !g_tts_prefetch_shutdown) {
        log_event("TTS_PRE", "Failed to create prefetch events");
        return;
    }

    g_tts_prefetch_thread = CreateThread(NULL, 0, tts_prefetch_proc, NULL, 0, NULL);
    if (!g_tts_prefetch_thread) {
        log_event("TTS_PRE", "Failed to create prefetch thread");
        return;
    }

    /* Kick off initial sweep */
    InterlockedExchange(&g_tts_prefetch_total, (LONG)g_tts_groupings_count);
    InterlockedExchange(&g_tts_prefetch_done, 0);
    SetEvent(g_tts_prefetch_event);

    log_event("TTS_PRE", "Prefetch thread started");
}

static void tts_prefetch_stop(void) {
    {
        HINTERNET h = InterlockedExchangePointer(
            (volatile PVOID *)&g_tts_prefetch_hrequest, NULL);
        if (h) WinHttpCloseHandle(h);
    }
    if (g_tts_prefetch_shutdown)
        SetEvent(g_tts_prefetch_shutdown);
    if (g_tts_prefetch_thread) {
        WaitForSingleObject(g_tts_prefetch_thread, 5000);
        CloseHandle(g_tts_prefetch_thread);
        g_tts_prefetch_thread = NULL;
    }
    if (g_tts_prefetch_event) {
        CloseHandle(g_tts_prefetch_event);
        g_tts_prefetch_event = NULL;
    }
    if (g_tts_prefetch_shutdown) {
        CloseHandle(g_tts_prefetch_shutdown);
        g_tts_prefetch_shutdown = NULL;
    }
}

/* Set priority sentence and wake prefetch thread. */
static void tts_prefetch_prioritize(int idx) {
    InterlockedExchange(&g_tts_prefetch_priority, (LONG)idx);
    if (g_tts_prefetch_event)
        SetEvent(g_tts_prefetch_event);
}

/* Publish grouping timestamps to g_tts_current_ts for drill renderer. */
static void tts_publish_cached_timestamps(int idx) {
    TtsTimestamps ts = {0};
    int have = tts_groupings_copy(idx, &ts);
    EnterCriticalSection(&g_tts_current_ts_lock);
    free(g_tts_current_ts.words);
    if (have) {
        g_tts_current_ts = ts;
    } else {
        g_tts_current_ts.words = NULL;
        g_tts_current_ts.count = 0;
    }
    LeaveCriticalSection(&g_tts_current_ts_lock);
}

/* Queue text for server TTS playback (non-blocking).
 * sentence_idx >= 0 stores word groupings on success; -1 skips.
 * seed: -1=use voice's locked seed (or random), -2=force random (tuning), >=0=explicit */
static void tts_speak_server(const char *text, int sentence_idx, int seed) {
    if (!text || !text[0]) return;

    /* Interrupt any in-progress playback */
    InterlockedExchange(&g_tts_interrupt, 1);

    char *copy = _strdup(text);
    if (!copy) return;

    EnterCriticalSection(&g_tts_lock);
    free(g_tts_pending_text);
    g_tts_pending_text = copy;
    g_tts_pending_voice_idx = g_tts_voice_idx;
    g_tts_pending_sentence_idx = sentence_idx;
    g_tts_pending_seed = seed;
    LeaveCriticalSection(&g_tts_lock);

    SetEvent(g_tts_request_event);
}

/* ---- Local LLM (llama-server HTTP) ---- */

static void llm_history_clear(void) {
    g_llm_history_count = 0;
}

static void llm_history_append(const char *role, const char *content) {
    if (g_llm_history_count >= LLM_MAX_HISTORY) {
        /* Drop oldest message (shift left by 1) */
        memmove(&g_llm_history[0], &g_llm_history[1],
                (LLM_MAX_HISTORY - 1) * sizeof(LlmMessage));
        g_llm_history_count = LLM_MAX_HISTORY - 1;
    }
    LlmMessage *m = &g_llm_history[g_llm_history_count++];
    strncpy(m->role, role, sizeof(m->role) - 1);
    m->role[sizeof(m->role) - 1] = '\0';
    strncpy(m->content, content, sizeof(m->content) - 1);
    m->content[sizeof(m->content) - 1] = '\0';
}

/* Escape a string for JSON: handle ", \, \n, \r, \t */
static int json_escape(const char *src, char *dst, int dst_size) {
    int j = 0;
    for (int i = 0; src[i] && j < dst_size - 2; i++) {
        char c = src[i];
        if (c == '"' || c == '\\') {
            if (j + 2 >= dst_size) break;
            dst[j++] = '\\'; dst[j++] = c;
        } else if (c == '\n') {
            if (j + 2 >= dst_size) break;
            dst[j++] = '\\'; dst[j++] = 'n';
        } else if (c == '\r') {
            if (j + 2 >= dst_size) break;
            dst[j++] = '\\'; dst[j++] = 'r';
        } else if (c == '\t') {
            if (j + 2 >= dst_size) break;
            dst[j++] = '\\'; dst[j++] = 't';
        } else {
            dst[j++] = c;
        }
    }
    dst[j] = '\0';
    return j;
}

/* Build OpenAI-compatible JSON request with conversation history */
static int llm_build_request_json(const char *prompt, char *buf, int size) {
    char escaped[LLM_MAX_CONTENT * 2];
    int pos = 0;

    const char *sys_prompt = g_tutor_mode ? g_tutor_system_prompt : g_llm_system_prompt;
    pos += snprintf(buf + pos, size - pos,
        "{\"model\":\"local\",\"messages\":[{\"role\":\"system\",\"content\":\"");
    json_escape(sys_prompt, escaped, sizeof(escaped));
    pos += snprintf(buf + pos, size - pos, "%s\"}", escaped);

    /* Append conversation history */
    for (int i = 0; i < g_llm_history_count && pos < size - 256; i++) {
        json_escape(g_llm_history[i].content, escaped, sizeof(escaped));
        pos += snprintf(buf + pos, size - pos,
            ",{\"role\":\"%s\",\"content\":\"%s\"}",
            g_llm_history[i].role, escaped);
    }

    /* Append current user prompt */
    json_escape(prompt, escaped, sizeof(escaped));
    pos += snprintf(buf + pos, size - pos,
        ",{\"role\":\"user\",\"content\":\"%s\"}", escaped);

    int max_tokens = g_tutor_mode ? 512 : 256;
    pos += snprintf(buf + pos, size - pos,
        "],\"max_tokens\":%d,\"temperature\":0.7}", max_tokens);

    return pos;
}

/* Extract content from OpenAI-compatible JSON response.
 * Looks for "content":" inside choices[0].message */
static int llm_parse_response(const char *json, char *buf, int size) {
    /* Find choices[0].message.content — search for "content":" after "message" */
    const char *msg = strstr(json, "\"message\"");
    if (!msg) return -1;
    const char *key = strstr(msg, "\"content\"");
    if (!key) return -1;
    const char *colon = strchr(key + 9, ':');
    if (!colon) return -1;
    /* Skip whitespace and opening quote */
    const char *p = colon + 1;
    while (*p == ' ' || *p == '\t' || *p == '\n' || *p == '\r') p++;
    if (*p == 'n' && strncmp(p, "null", 4) == 0) {
        buf[0] = '\0';
        return 0;
    }
    if (*p != '"') return -1;
    p++; /* skip opening quote */

    /* Copy content, handling JSON escape sequences */
    int j = 0;
    while (*p && *p != '"' && j < size - 1) {
        if (*p == '\\' && *(p + 1)) {
            p++;
            switch (*p) {
                case '"':  buf[j++] = '"'; break;
                case '\\': buf[j++] = '\\'; break;
                case 'n':  buf[j++] = '\n'; break;
                case 'r':  buf[j++] = '\r'; break;
                case 't':  buf[j++] = '\t'; break;
                case '/':  buf[j++] = '/'; break;
                default:   buf[j++] = *p; break;
            }
        } else {
            buf[j++] = *p;
        }
        p++;
    }
    buf[j] = '\0';
    return j;
}

/* LLM worker thread: wait for prompts, POST to llama-server, return response */
static DWORD WINAPI llm_worker_proc(LPVOID param) {
    (void)param;
    HANDLE events[2] = { g_llm_request_event, g_llm_shutdown_event };

    HINTERNET hSession = WinHttpOpen(L"VoiceNoteGUI/1.0",
                                      WINHTTP_ACCESS_TYPE_NO_PROXY,
                                      WINHTTP_NO_PROXY_NAME,
                                      WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) {
        log_event("LLM", "WinHttpOpen failed");
        return 1;
    }

    while (1) {
        DWORD which = WaitForMultipleObjects(2, events, FALSE, INFINITE);
        if (which != WAIT_OBJECT_0) break; /* shutdown */

        /* Grab pending prompt */
        EnterCriticalSection(&g_llm_lock);
        char *prompt = g_llm_pending_prompt;
        g_llm_pending_prompt = NULL;
        LeaveCriticalSection(&g_llm_lock);

        if (!prompt) continue;

        /* Check interrupt before starting */
        InterlockedExchange(&g_llm_interrupt, 0);
        if (InterlockedCompareExchange(&g_llm_interrupt, 0, 0)) {
            free(prompt);
            continue;
        }

        log_event("LLM", "Sending request...");

        /* Build JSON request body */
        char *request_buf = (char *)malloc(LLM_REQUEST_BUF);
        if (!request_buf) { free(prompt); continue; }
        llm_build_request_json(prompt, request_buf, LLM_REQUEST_BUF);

        /* Connect to llama-server */
        HINTERNET hConnect = WinHttpConnect(hSession, L"localhost",
                                             LLM_SERVER_PORT, 0);
        if (!hConnect) {
            log_event("LLM", "WinHttpConnect failed");
            g_llm_server_ok = 0;
            free(request_buf);
            free(prompt);
            continue;
        }

        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
                                                 L"/v1/chat/completions",
                                                 NULL, WINHTTP_NO_REFERER,
                                                 WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
        if (!hRequest) {
            log_event("LLM", "WinHttpOpenRequest failed");
            g_llm_server_ok = 0;
            WinHttpCloseHandle(hConnect);
            free(request_buf);
            free(prompt);
            continue;
        }

        /* Set timeouts: 2s connect, 30s receive */
        WinHttpSetTimeouts(hRequest, 2000, 2000, 30000, 30000);

        /* Send request */
        DWORD req_len = (DWORD)strlen(request_buf);
        BOOL ok = WinHttpSendRequest(hRequest,
                                      L"Content-Type: application/json\r\n",
                                      (DWORD)-1L,
                                      request_buf, req_len, req_len, 0);

        if (ok) ok = WinHttpReceiveResponse(hRequest, NULL);

        char *response_buf = NULL;
        int response_len = 0;

        if (ok) {
            /* Check if interrupted while waiting */
            if (InterlockedCompareExchange(&g_llm_interrupt, 0, 0)) {
                log_event("LLM", "Interrupted during receive");
                WinHttpCloseHandle(hRequest);
                WinHttpCloseHandle(hConnect);
                free(request_buf);
                free(prompt);
                continue;
            }

            response_buf = (char *)malloc(LLM_RESPONSE_BUF);
            if (response_buf) {
                DWORD bytes_read = 0;
                DWORD total = 0;
                while (WinHttpReadData(hRequest, response_buf + total,
                                        LLM_RESPONSE_BUF - total - 1, &bytes_read)) {
                    if (bytes_read == 0) break;
                    total += bytes_read;
                    if (total >= (DWORD)(LLM_RESPONSE_BUF - 1)) break;
                }
                response_buf[total] = '\0';
                response_len = (int)total;
            }
            g_llm_server_ok = 1;
        } else {
            log_event("LLM", "HTTP request failed");
            g_llm_server_ok = 0;
        }

        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        free(request_buf);

        /* Parse response and post to GUI */
        if (response_buf && response_len > 0) {
            char content[LLM_MAX_CONTENT];
            if (llm_parse_response(response_buf, content, sizeof(content)) > 0) {
                /* Update history: add user prompt + assistant response */
                llm_history_append("user", prompt);
                llm_history_append("assistant", content);

                char *copy = _strdup(content);
                if (copy) {
                    PostMessageA(g_hwnd_main, WM_LLM_RESPONSE, 0, (LPARAM)copy);
                }

                char msg[64];
                snprintf(msg, sizeof(msg), "Response: %d chars", (int)strlen(content));
                log_event("LLM", msg);
            } else {
                log_event("LLM", "Failed to parse response");
                /* Log first 200 chars of raw response for debugging */
                char dbg[204];
                strncpy(dbg, response_buf, 200);
                dbg[200] = '\0';
                log_event("LLM_RAW", dbg);
            }
            free(response_buf);
        }

        free(prompt);
    }

    WinHttpCloseHandle(hSession);
    return 0;
}

static int llm_worker_start(void) {
    InitializeCriticalSection(&g_llm_lock);
    g_llm_request_event = CreateEventA(NULL, FALSE, FALSE, NULL);   /* auto-reset */
    g_llm_shutdown_event = CreateEventA(NULL, TRUE, FALSE, NULL);   /* manual-reset */
    if (!g_llm_request_event || !g_llm_shutdown_event) {
        log_event("LLM", "Failed to create worker events");
        return 0;
    }

    g_llm_thread = CreateThread(NULL, 0, llm_worker_proc, NULL, 0, NULL);
    if (!g_llm_thread) {
        log_event("LLM", "Failed to create worker thread");
        return 0;
    }

    log_event("LLM", "Worker thread started");
    return 1;
}

static void llm_worker_stop(void) {
    if (g_llm_shutdown_event) {
        SetEvent(g_llm_shutdown_event);
    }
    if (g_llm_thread) {
        WaitForSingleObject(g_llm_thread, 3000);
        CloseHandle(g_llm_thread);
        g_llm_thread = NULL;
    }
    if (g_llm_request_event) {
        CloseHandle(g_llm_request_event);
        g_llm_request_event = NULL;
    }
    if (g_llm_shutdown_event) {
        CloseHandle(g_llm_shutdown_event);
        g_llm_shutdown_event = NULL;
    }
    EnterCriticalSection(&g_llm_lock);
    free(g_llm_pending_prompt);
    g_llm_pending_prompt = NULL;
    LeaveCriticalSection(&g_llm_lock);
    DeleteCriticalSection(&g_llm_lock);
}

/* Queue a prompt for the LLM worker thread (non-blocking) */
static void llm_send(const char *text) {
    if (!text || !text[0]) return;

    /* Interrupt any pending request */
    InterlockedExchange(&g_llm_interrupt, 1);

    char *copy = _strdup(text);
    if (!copy) return;

    EnterCriticalSection(&g_llm_lock);
    free(g_llm_pending_prompt);
    g_llm_pending_prompt = copy;
    LeaveCriticalSection(&g_llm_lock);

    SetEvent(g_llm_request_event);
}

/* Strip CJK characters from UTF-8 text (keep ASCII, Latin, pinyin tone marks).
 * Removes codepoints in U+3000..U+9FFF (CJK Unified + symbols) and
 * U+FF00..U+FF60 (fullwidth forms). Collapses whitespace. Returns malloc'd string. */
static char *strip_cjk(const char *text) {
    if (!text) return NULL;
    int len = (int)strlen(text);
    char *out = (char *)malloc(len + 1);
    if (!out) return NULL;

    int j = 0;
    const unsigned char *p = (const unsigned char *)text;
    while (*p) {
        unsigned int cp = 0;
        int bytes = 0;

        if (*p < 0x80)                { cp = *p;          bytes = 1; }
        else if ((*p & 0xE0) == 0xC0) { cp = *p & 0x1F;   bytes = 2; }
        else if ((*p & 0xF0) == 0xE0) { cp = *p & 0x0F;   bytes = 3; }
        else if ((*p & 0xF8) == 0xF0) { cp = *p & 0x07;   bytes = 4; }
        else { p++; continue; }

        int valid = 1;
        for (int i = 1; i < bytes; i++) {
            if ((p[i] & 0xC0) != 0x80) { valid = 0; break; }
            cp = (cp << 6) | (p[i] & 0x3F);
        }
        if (!valid) { p++; continue; }

        int is_cjk = (cp >= 0x3000 && cp <= 0x9FFF) ||
                     (cp >= 0xFF00 && cp <= 0xFF60);
        if (!is_cjk) {
            for (int i = 0; i < bytes; i++)
                out[j++] = (char)p[i];
        }
        p += bytes;
    }
    out[j] = '\0';

    /* Collapse runs of spaces/tabs */
    int src = 0, dst = 0, prev_space = 0;
    while (out[src]) {
        if (out[src] == ' ' || out[src] == '\t') {
            if (!prev_space) out[dst++] = ' ';
            prev_space = 1;
        } else {
            out[dst++] = out[src];
            prev_space = 0;
        }
        src++;
    }
    out[dst] = '\0';
    return out;
}

/* Remove tutor section labels (Chinese/Pinyin/English/Grammar/Prompt) from text.
 * Labels with content have the label removed; orphaned labels (no content after
 * CJK stripping) are removed entirely. Modifies string in-place. */
static void strip_tutor_labels(char *text) {
    if (!text) return;
    int len = (int)strlen(text);
    char *out = (char *)malloc(len + 1);
    if (!out) return;

    static const char *labels[] = { "Chinese", "Pinyin", "English", "Grammar", "Prompt" };
    static const int label_lens[] = { 7, 6, 7, 7, 6 };
    #define NUM_LABELS 5

    int ci = 0, co = 0;
    while (text[ci]) {
        /* Check if we're at a known label */
        int matched = 0;
        for (int k = 0; k < NUM_LABELS; k++) {
            if (strncmp(text + ci, labels[k], label_lens[k]) == 0 &&
                text[ci + label_lens[k]] == ':') {
                int after = ci + label_lens[k] + 1;
                while (text[after] == ' ') after++;
                /* Check if content follows or it's orphaned (only whitespace/next label/end) */
                int peek = after;
                while (text[peek] == '\n' || text[peek] == '\r') peek++;
                int nl = peek;
                while (text[nl] && ((text[nl] >= 'A' && text[nl] <= 'Z') ||
                                    (text[nl] >= 'a' && text[nl] <= 'z'))) nl++;
                if (text[peek] == '\0' || (nl > peek && text[nl] == ':')) {
                    ci = peek; /* Orphaned — skip entirely */
                } else {
                    ci = after; /* Has content — skip label only */
                }
                matched = 1;
                break;
            }
        }
        if (!matched)
            out[co++] = text[ci++];
    }
    out[co] = '\0';

    /* Trim leading/trailing whitespace */
    int start = 0;
    while (out[start] == ' ' || out[start] == '\n' ||
           out[start] == '\r' || out[start] == '\t') start++;
    int end = co;
    while (end > start && (out[end - 1] == ' ' || out[end - 1] == '\n' ||
           out[end - 1] == '\r' || out[end - 1] == '\t')) end--;

    memcpy(text, out + start, end - start);
    text[end - start] = '\0';
    free(out);
    #undef NUM_LABELS
}

/* TTS: speak text asynchronously (non-blocking) */
static void tts_speak(const char *text) {
    if (!g_tts_enabled || !text || !text[0]) return;

    /* In tutor mode, strip CJK characters and section labels for TTS */
    const char *speak_text = text;
    char *stripped = NULL;
    if (g_tutor_mode) {
        stripped = strip_cjk(text);
        if (stripped) strip_tutor_labels(stripped);
        if (stripped && stripped[0]) {
            speak_text = stripped;
        } else {
            free(stripped);
            return;
        }
    }

    /* SAPI TTS */
    if (!g_tts_voice) { free(stripped); return; }

    /* Convert UTF-8 to wide string */
    int wlen = MultiByteToWideChar(CP_UTF8, 0, speak_text, -1, NULL, 0);
    if (wlen <= 0) { free(stripped); return; }
    wchar_t *wtext = (wchar_t *)malloc(wlen * sizeof(wchar_t));
    if (!wtext) { free(stripped); return; }
    MultiByteToWideChar(CP_UTF8, 0, speak_text, -1, wtext, wlen);

    /* Purge any current speech and queue new text */
    ISpVoice_Speak(g_tts_voice, wtext, SPF_ASYNC | SPF_PURGEBEFORESPEAK, NULL);
    free(wtext);
    free(stripped);
}

/* Conversation history: append a message to the chat log */
static void chat_append(const char *role, const char *text) {
    if (!text || !text[0]) return;

    /* Format: "[role] text\r\n" */
    char line[PIPE_BUF_SIZE + 64];
    snprintf(line, sizeof(line), "[%s] %s\r\n", role, text);
    int line_len = (int)strlen(line);

    if (g_chat_len + line_len < MAX_CHAT_LEN - 1) {
        memcpy(g_chat_log + g_chat_len, line, line_len);
        g_chat_len += line_len;
        g_chat_log[g_chat_len] = '\0';
    }

    if (g_hwnd_chat) {
        SetWindowTextUTF8(g_hwnd_chat, g_chat_log);
        /* Scroll to bottom — use wide length for correct caret position */
        int wlen = MultiByteToWideChar(CP_UTF8, 0, g_chat_log, g_chat_len, NULL, 0);
        SendMessageW(g_hwnd_chat, EM_SETSEL, (WPARAM)wlen, (LPARAM)wlen);
        SendMessageW(g_hwnd_chat, EM_SCROLLCARET, 0, 0);
    }
}

/* Named pipe server thread - creates pipe, waits for client, reads responses.
 * Uses overlapped I/O so ReadFile here doesn't block WriteFile from the GUI thread. */
static DWORD WINAPI pipe_thread_proc(LPVOID param) {
    (void)param;

    while (g_pipe_running) {
        /* Create named pipe with FILE_FLAG_OVERLAPPED for async I/O */
        g_pipe = CreateNamedPipeA(
            PIPE_NAME,
            PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            1,               /* max instances */
            PIPE_BUF_SIZE,   /* out buffer */
            PIPE_BUF_SIZE,   /* in buffer */
            0,               /* default timeout */
            NULL);           /* default security */

        if (g_pipe == INVALID_HANDLE_VALUE) {
            log_event("PIPE", "Failed to create named pipe");
            return 1;
        }

        log_event("PIPE", "Waiting for client connection...");

        /* Overlapped ConnectNamedPipe - wait for client or shutdown */
        OVERLAPPED ov_conn = {0};
        ov_conn.hEvent = CreateEventA(NULL, TRUE, FALSE, NULL);
        BOOL conn_ok = ConnectNamedPipe(g_pipe, &ov_conn);
        if (!conn_ok) {
            DWORD err = GetLastError();
            if (err == ERROR_IO_PENDING) {
                /* Wait for connection or shutdown */
                HANDLE waits[2] = { ov_conn.hEvent, g_pipe_shutdown_event };
                DWORD which = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
                if (which != WAIT_OBJECT_0) {
                    /* Shutdown signaled */
                    CancelIo(g_pipe);
                    CloseHandle(ov_conn.hEvent);
                    CloseHandle(g_pipe);
                    g_pipe = INVALID_HANDLE_VALUE;
                    break;
                }
            } else if (err != ERROR_PIPE_CONNECTED) {
                CloseHandle(ov_conn.hEvent);
                CloseHandle(g_pipe);
                g_pipe = INVALID_HANDLE_VALUE;
                break;
            }
        }
        CloseHandle(ov_conn.hEvent);

        if (!g_pipe_running) {
            CloseHandle(g_pipe);
            g_pipe = INVALID_HANDLE_VALUE;
            break;
        }

        g_pipe_connected = 1;
        log_event("PIPE", "Client connected");
        if (g_hwnd_stats) InvalidateRect(g_hwnd_stats, NULL, FALSE);

        /* Read responses from client using overlapped I/O */
        OVERLAPPED ov_read = {0};
        ov_read.hEvent = CreateEventA(NULL, TRUE, FALSE, NULL);
        char buf[PIPE_BUF_SIZE];

        while (g_pipe_running) {
            DWORD bytes_read = 0;
            ResetEvent(ov_read.hEvent);
            BOOL ok = ReadFile(g_pipe, buf, sizeof(buf) - 1, &bytes_read, &ov_read);

            if (!ok) {
                DWORD err = GetLastError();
                if (err == ERROR_IO_PENDING) {
                    /* Wait for data or shutdown */
                    HANDLE waits[2] = { ov_read.hEvent, g_pipe_shutdown_event };
                    DWORD which = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
                    if (which != WAIT_OBJECT_0) {
                        CancelIo(g_pipe);
                        break;
                    }
                    if (!GetOverlappedResult(g_pipe, &ov_read, &bytes_read, FALSE)) {
                        break;
                    }
                } else {
                    if (err == ERROR_BROKEN_PIPE || err == ERROR_PIPE_NOT_CONNECTED) {
                        log_event("PIPE", "Client disconnected");
                    }
                    break;
                }
            }

            if (bytes_read == 0) break;
            buf[bytes_read] = '\0';

            char *copy = _strdup(buf);
            if (copy) {
                PostMessageA(g_hwnd_main, WM_PIPE_RESPONSE, 0, (LPARAM)copy);
            }
        }

        CloseHandle(ov_read.hEvent);

        /* Client disconnected - clean up and loop to accept new client */
        g_pipe_connected = 0;
        DisconnectNamedPipe(g_pipe);
        CloseHandle(g_pipe);
        g_pipe = INVALID_HANDLE_VALUE;
        if (g_hwnd_stats) InvalidateRect(g_hwnd_stats, NULL, FALSE);
        log_event("PIPE", "Pipe reset, waiting for new client");
    }

    return 0;
}

/* Send text to connected pipe client (called from GUI thread).
 * Uses overlapped WriteFile so it doesn't block on the pipe thread's ReadFile. */
static void pipe_send(const char *text) {
    if (!g_pipe_connected || g_pipe == INVALID_HANDLE_VALUE) return;

    OVERLAPPED ov = {0};
    ov.hEvent = CreateEventA(NULL, TRUE, FALSE, NULL);
    DWORD bytes_written = 0;
    DWORD len = (DWORD)strlen(text);

    BOOL ok = WriteFile(g_pipe, text, len, &bytes_written, &ov);
    if (!ok && GetLastError() == ERROR_IO_PENDING) {
        /* Wait up to 1 second for write to complete */
        if (WaitForSingleObject(ov.hEvent, 1000) == WAIT_OBJECT_0) {
            GetOverlappedResult(g_pipe, &ov, &bytes_written, FALSE);
            ok = TRUE;
        } else {
            CancelIo(g_pipe);
            log_event("PIPE", "WriteFile timed out");
        }
    }

    if (ok) {
        char msg[64];
        snprintf(msg, sizeof(msg), "Sent %lu bytes", (unsigned long)bytes_written);
        log_event("PIPE", msg);
    } else if (GetLastError() != ERROR_IO_PENDING) {
        log_event("PIPE", "WriteFile failed");
    }

    CloseHandle(ov.hEvent);
}

/* Audio buffer functions */
static void add_audio_samples(const int16_t *pcm16, int sample_count) {
    EnterCriticalSection(&g_audio_lock);

    if (!g_capture_ready && sample_count > 0)
        g_capture_ready = 1;

    float energy = 0.0f;
    for (int i = 0; i < sample_count && g_audio_samples < MAX_AUDIO_SAMPLES; i++) {
        float sample = pcm16[i] / 32768.0f;
        g_audio_buffer[g_audio_write_pos] = sample;
        g_audio_write_pos = (g_audio_write_pos + 1) % MAX_AUDIO_SAMPLES;
        if (g_audio_samples < MAX_AUDIO_SAMPLES) {
            g_audio_samples++;
        }
        /* Also save to full recording buffer (never cleared during recording) */
        if (g_recording_samples < MAX_AUDIO_SAMPLES) {
            g_recording_buffer[g_recording_samples++] = sample;
        }
        energy += fabsf(sample);
    }

    if (sample_count > 0) {
        g_current_energy = energy / sample_count;
    }

    LeaveCriticalSection(&g_audio_lock);
}

static int get_audio_samples(float *dest, int max_samples) {
    EnterCriticalSection(&g_audio_lock);
    int count = g_audio_samples < max_samples ? g_audio_samples : max_samples;
    int start = (g_audio_write_pos - count + MAX_AUDIO_SAMPLES) % MAX_AUDIO_SAMPLES;
    for (int i = 0; i < count; i++) {
        dest[i] = g_audio_buffer[(start + i) % MAX_AUDIO_SAMPLES];
    }
    LeaveCriticalSection(&g_audio_lock);
    return count;
}

static void clear_audio_buffer(void) {
    EnterCriticalSection(&g_audio_lock);
    g_audio_write_pos = 0;
    g_audio_samples = 0;
    g_current_energy = 0.0f;
    LeaveCriticalSection(&g_audio_lock);
}

/* Forward declarations */
static void update_scrollbar(void);

/* Write audio buffer as WAV file for offline testing */
static void write_wav(const float *samples, int n_samples) {
    /* Create recordings directory next to executable */
    CreateDirectoryA("recordings", NULL);

    /* Generate filename with timestamp */
    time_t now = time(NULL);
    struct tm *tm_info = localtime(&now);
    char filename[MAX_PATH];
    snprintf(filename, sizeof(filename), "recordings\\recording_%04d%02d%02d_%02d%02d%02d.wav",
             tm_info->tm_year + 1900, tm_info->tm_mon + 1, tm_info->tm_mday,
             tm_info->tm_hour, tm_info->tm_min, tm_info->tm_sec);

    FILE *f = fopen(filename, "wb");
    if (!f) {
        log_event("WAV_ERR", "Failed to open WAV file for writing");
        return;
    }

    /* WAV header - 16-bit PCM, 16kHz, mono */
    int32_t data_size = n_samples * 2;  /* 16-bit = 2 bytes per sample */
    int32_t file_size = 36 + data_size;
    int16_t audio_format = 1;       /* PCM */
    int16_t num_channels = 1;       /* mono */
    int32_t sample_rate = WHISPER_SAMPLE_RATE;
    int32_t byte_rate = sample_rate * 2;
    int16_t block_align = 2;
    int16_t bits_per_sample = 16;

    /* RIFF header */
    fwrite("RIFF", 1, 4, f);
    fwrite(&file_size, 4, 1, f);
    fwrite("WAVE", 1, 4, f);

    /* fmt chunk */
    fwrite("fmt ", 1, 4, f);
    int32_t fmt_size = 16;
    fwrite(&fmt_size, 4, 1, f);
    fwrite(&audio_format, 2, 1, f);
    fwrite(&num_channels, 2, 1, f);
    fwrite(&sample_rate, 4, 1, f);
    fwrite(&byte_rate, 4, 1, f);
    fwrite(&block_align, 2, 1, f);
    fwrite(&bits_per_sample, 2, 1, f);

    /* data chunk */
    fwrite("data", 1, 4, f);
    fwrite(&data_size, 4, 1, f);

    /* Convert float samples to int16 and write */
    for (int i = 0; i < n_samples; i++) {
        float s = samples[i];
        if (s > 1.0f) s = 1.0f;
        if (s < -1.0f) s = -1.0f;
        int16_t pcm = (int16_t)(s * 32767.0f);
        fwrite(&pcm, 2, 1, f);
    }

    fclose(f);

    char msg[MAX_PATH + 32];
    snprintf(msg, sizeof(msg), "Saved %s (%d samples, %.1fs)", filename, n_samples,
             (float)n_samples / WHISPER_SAMPLE_RATE);
    log_event("WAV_SAVE", msg);
}

/* Helper: convert bar index to time and vice versa */
static float bar_to_time(int bar) {
    return (float)bar * WAVEFORM_UPDATE_MS / 1000.0f;
}

static int time_to_bar(float time_sec) {
    return (int)(time_sec * 1000.0f / WAVEFORM_UPDATE_MS);
}

/* Helper: convert x position to bar index (in stopped mode) */
static int x_to_bar(int x, int width) {
    int bar_width = (width - 20) / WAVEFORM_BARS;
    int bar = (x - 10) / bar_width + g_scroll_offset;
    if (bar < 0) bar = 0;
    if (bar >= g_stored_bar_count) bar = g_stored_bar_count - 1;
    return bar;
}

/* Waveform window procedure - shows live audio input OR stored waveform when stopped */
static LRESULT CALLBACK WaveformWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);

            RECT rc;
            GetClientRect(hwnd, &rc);
            int width = rc.right - rc.left;
            int height = rc.bottom - rc.top;

            /* Double buffer */
            HDC memdc = CreateCompatibleDC(hdc);
            HBITMAP membmp = CreateCompatibleBitmap(hdc, width, height);
            HBITMAP oldbmp = (HBITMAP)SelectObject(memdc, membmp);

            /* Background */
            HBRUSH bgbrush = CreateSolidBrush(COLOR_BG);
            FillRect(memdc, &rc, bgbrush);
            DeleteObject(bgbrush);

            /* Draw bars */
            int bar_width = (width - 20) / WAVEFORM_BARS;
            int bar_gap = 2;
            int max_bar_height = height - 16;

            /* Draw all bars - use stored data when stopped, live data when recording */
            for (int i = 0; i < WAVEFORM_BARS; i++) {
                float level;
                if (!g_is_recording && g_stored_bar_count > 0) {
                    /* Stopped mode: show stored waveform at scroll offset */
                    int stored_idx = g_scroll_offset + i;
                    if (stored_idx >= 0 && stored_idx < g_stored_bar_count) {
                        level = g_stored_levels[stored_idx];
                    } else {
                        level = 0.0f;
                    }
                } else {
                    /* Recording mode: show live data */
                    level = g_waveform_levels[i];
                }

                int bar_height;
                if (g_is_recording && !g_capture_ready) {
                    /* Warmup: show low static bars so orange is visible */
                    bar_height = max_bar_height / 8;
                } else {
                    bar_height = (int)(level * max_bar_height);
                }
                if (bar_height < 2) bar_height = 2;
                if (bar_height > max_bar_height) bar_height = max_bar_height;

                int x = 10 + i * bar_width;
                int y = height - 8 - bar_height;

                /* Bright colors for live audio */
                COLORREF color;
                if (g_is_recording && !g_capture_ready) {
                    /* Warming up: orange bars until audio flows */
                    color = RGB(220, 140, 40);
                } else if (level < 0.3f) {
                    color = COLOR_WAVE_LOW;
                } else if (level < 0.6f) {
                    color = COLOR_WAVE_MED;
                } else {
                    color = COLOR_WAVE_HIGH;
                }

                /* Tint bars by transcription state during recording */
                if (g_is_recording && g_committed_samples > 0) {
                    int global_bar = g_stored_bar_count - WAVEFORM_BARS + i;
                    int committed_bar = g_committed_samples / SAMPLES_PER_BAR;
                    int window_end_bar = (g_committed_samples + g_window_samples) / SAMPLES_PER_BAR;
                    int r = GetRValue(color);
                    int g = GetGValue(color);
                    int b = GetBValue(color);
                    if (global_bar < committed_bar) {
                        /* Committed: reduce R and B by 40% -- appears greener */
                        r = r * 60 / 100;
                        b = b * 60 / 100;
                        color = RGB(r, g, b);
                    } else if (global_bar < window_end_bar) {
                        /* WIP window: reduce B by 50% -- appears yellower */
                        b = b * 50 / 100;
                        color = RGB(r, g, b);
                    }
                }

                HBRUSH brush = CreateSolidBrush(color);
                RECT bar_rc = { x, y, x + bar_width - bar_gap, height - 8 };
                FillRect(memdc, &bar_rc, brush);
                DeleteObject(brush);
            }

            if (g_is_recording) {
                /* Silence threshold line (only during recording) */
                int threshold_y = height - 8 - (int)(SILENCE_THRESHOLD * 20.0f * max_bar_height);
                if (threshold_y > 8 && threshold_y < height - 8) {
                    HPEN pen = CreatePen(PS_DASH, 1, COLOR_SILENCE);
                    HPEN oldpen = (HPEN)SelectObject(memdc, pen);
                    MoveToEx(memdc, 10, threshold_y, NULL);
                    LineTo(memdc, width - 10, threshold_y);
                    SelectObject(memdc, oldpen);
                    DeleteObject(pen);
                }

                /* Commit boundary line (green vertical line at committed position) */
                if (g_committed_samples > 0) {
                    int committed_bar = g_committed_samples / SAMPLES_PER_BAR;
                    int live_bar = committed_bar - (g_stored_bar_count - WAVEFORM_BARS);
                    if (live_bar >= 0 && live_bar < WAVEFORM_BARS) {
                        int cx = 10 + live_bar * bar_width;
                        HPEN cpen = CreatePen(PS_SOLID, 2, COLOR_WAVE_LOW);
                        HPEN oldcpen = (HPEN)SelectObject(memdc, cpen);
                        MoveToEx(memdc, cx, 4, NULL);
                        LineTo(memdc, cx, height - 4);
                        SelectObject(memdc, oldcpen);
                        DeleteObject(cpen);
                    }
                }
            } else {
                /* Draw marker line when stopped */
                if (g_marker_bar >= 0 && g_marker_bar >= g_scroll_offset &&
                    g_marker_bar < g_scroll_offset + WAVEFORM_BARS) {
                    int marker_x = 10 + (g_marker_bar - g_scroll_offset) * bar_width + bar_width / 2;
                    HPEN pen = CreatePen(PS_SOLID, 2, RGB(255, 100, 100));
                    HPEN oldpen = (HPEN)SelectObject(memdc, pen);
                    MoveToEx(memdc, marker_x, 4, NULL);
                    LineTo(memdc, marker_x, height - 4);
                    SelectObject(memdc, oldpen);
                    DeleteObject(pen);

                    /* Draw time at marker */
                    char time_str[32];
                    snprintf(time_str, sizeof(time_str), "%.1fs", g_marker_time);
                    SetBkMode(memdc, TRANSPARENT);
                    SetTextColor(memdc, RGB(255, 100, 100));
                    SelectObject(memdc, g_font_small);
                    TextOutA(memdc, marker_x + 4, 4, time_str, (int)strlen(time_str));
                }
            }

            /* Blit */
            BitBlt(hdc, 0, 0, width, height, memdc, 0, 0, SRCCOPY);

            SelectObject(memdc, oldbmp);
            DeleteObject(membmp);
            DeleteDC(memdc);

            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_LBUTTONDOWN:
            if (!g_is_recording && g_stored_bar_count > 0) {
                RECT rc;
                GetClientRect(hwnd, &rc);
                int x = LOWORD(lParam);
                g_marker_bar = x_to_bar(x, rc.right - rc.left);
                g_marker_time = bar_to_time(g_marker_bar);
                g_dragging = 1;
                SetCapture(hwnd);
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;

        case WM_MOUSEMOVE:
            if (g_dragging && !g_is_recording) {
                RECT rc;
                GetClientRect(hwnd, &rc);
                int x = LOWORD(lParam);
                g_marker_bar = x_to_bar(x, rc.right - rc.left);
                g_marker_time = bar_to_time(g_marker_bar);
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;

        case WM_LBUTTONUP:
            if (g_dragging) {
                g_dragging = 0;
                ReleaseCapture();
            }
            return 0;

        case WM_MOUSEWHEEL:
            if (!g_is_recording && g_stored_bar_count > WAVEFORM_BARS) {
                int delta = GET_WHEEL_DELTA_WPARAM(wParam);
                int scroll_amount = (delta > 0) ? -5 : 5;  /* Scroll 5 bars at a time */
                int new_offset = g_scroll_offset + scroll_amount;
                int max_offset = g_stored_bar_count - WAVEFORM_BARS;
                if (new_offset < 0) new_offset = 0;
                if (new_offset > max_offset) new_offset = max_offset;
                if (new_offset != g_scroll_offset) {
                    g_scroll_offset = new_offset;
                    update_scrollbar();
                    InvalidateRect(g_hwnd_waveform, NULL, FALSE);
                }
            }
            return 0;

        case WM_ERASEBKGND:
            return 1;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* Query process CPU% and working set */
static void query_process_resources(void) {
    /* Working set */
    PROCESS_MEMORY_COUNTERS pmc;
    if (GetProcessMemoryInfo(GetCurrentProcess(), &pmc, sizeof(pmc))) {
        g_working_set_mb = pmc.WorkingSetSize / (1024 * 1024);
    }

    /* CPU% via GetProcessTimes delta */
    FILETIME ft_create, ft_exit, ft_kernel, ft_user;
    if (GetProcessTimes(GetCurrentProcess(), &ft_create, &ft_exit, &ft_kernel, &ft_user)) {
        ULONGLONG kernel = ((ULONGLONG)ft_kernel.dwHighDateTime << 32) | ft_kernel.dwLowDateTime;
        ULONGLONG user = ((ULONGLONG)ft_user.dwHighDateTime << 32) | ft_user.dwLowDateTime;
        ULONGLONG now = GetTickCount64() * 10000ULL;  /* ms to 100ns units */

        if (g_cpu_prev_time > 0) {
            ULONGLONG dt = now - g_cpu_prev_time;
            if (dt > 0) {
                ULONGLONG dk = kernel - g_cpu_prev_kernel;
                ULONGLONG du = user - g_cpu_prev_user;
                g_cpu_percent = (double)(dk + du) * 100.0 / (double)dt;
            }
        }
        g_cpu_prev_kernel = kernel;
        g_cpu_prev_user = user;
        g_cpu_prev_time = now;
    }
}

/* Query system/hardware info (called once at startup) */
static void query_system_info(void) {
    /* OS version: map to friendly name (Windows 11, Windows 10, etc.) */
    {
        OSVERSIONINFOW ovi = { .dwOSVersionInfoSize = sizeof(ovi) };
        typedef LONG (WINAPI *RtlGetVersion_t)(OSVERSIONINFOW *);
        RtlGetVersion_t fn = (RtlGetVersion_t)GetProcAddress(
            GetModuleHandleW(L"ntdll.dll"), "RtlGetVersion");
        if (fn && fn(&ovi) == 0) {
            const char *name = "Windows";
            if (ovi.dwMajorVersion == 10 && ovi.dwBuildNumber >= 22000)
                name = "Windows 11";
            else if (ovi.dwMajorVersion == 10)
                name = "Windows 10";
            else if (ovi.dwMajorVersion == 6 && ovi.dwMinorVersion == 3)
                name = "Windows 8.1";
            else if (ovi.dwMajorVersion == 6 && ovi.dwMinorVersion == 2)
                name = "Windows 8";
            else if (ovi.dwMajorVersion == 6 && ovi.dwMinorVersion == 1)
                name = "Windows 7";
            snprintf(g_os_version, sizeof(g_os_version), "%s", name);
        } else {
            snprintf(g_os_version, sizeof(g_os_version), "Windows");
        }
    }

    /* CPU name from registry, truncated at " with " (removes iGPU suffix) */
    {
        HKEY hkey;
        if (RegOpenKeyExA(HKEY_LOCAL_MACHINE,
                "HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0",
                0, KEY_READ, &hkey) == ERROR_SUCCESS) {
            DWORD size = sizeof(g_cpu_name);
            DWORD type = 0;
            if (RegQueryValueExA(hkey, "ProcessorNameString", NULL, &type,
                    (BYTE *)g_cpu_name, &size) != ERROR_SUCCESS) {
                snprintf(g_cpu_name, sizeof(g_cpu_name), "Unknown CPU");
            }
            RegCloseKey(hkey);
            /* Trim leading spaces */
            char *p = g_cpu_name;
            while (*p == ' ') p++;
            if (p != g_cpu_name)
                memmove(g_cpu_name, p, strlen(p) + 1);
            /* Truncate at " with " (e.g. "... with Radeon Graphics") */
            char *with = strstr(g_cpu_name, " with ");
            if (with) *with = '\0';
        } else {
            snprintf(g_cpu_name, sizeof(g_cpu_name), "Unknown CPU");
        }
    }

    /* Total RAM (installed, from SMBIOS via GetPhysicallyInstalledSystemMemory) */
    {
        ULONGLONG kb = 0;
        if (GetPhysicallyInstalledSystemMemory(&kb)) {
            double gb = (double)kb / (1024.0 * 1024.0);
            snprintf(g_ram_total, sizeof(g_ram_total), "%.0f GB", gb);
        } else {
            /* Fallback to usable memory */
            MEMORYSTATUSEX ms = { .dwLength = sizeof(ms) };
            if (GlobalMemoryStatusEx(&ms)) {
                double gb = (double)ms.ullTotalPhys / (1024.0 * 1024.0 * 1024.0);
                snprintf(g_ram_total, sizeof(g_ram_total), "%.0f GB", gb);
            } else {
                snprintf(g_ram_total, sizeof(g_ram_total), "?");
            }
        }
    }

    /* Logical cores */
    {
        SYSTEM_INFO si;
        GetSystemInfo(&si);
        snprintf(g_cpu_cores, sizeof(g_cpu_cores), "%lu", si.dwNumberOfProcessors);
    }
}

/* Query OS microphone and speaker volume/mute state */
static void query_device_status(void) {
    IMMDeviceEnumerator *enumerator = NULL;
    HRESULT hr = CoCreateInstance(
        &CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
        &IID_IMMDeviceEnumerator, (void **)&enumerator);
    if (FAILED(hr)) return;

    /* Query microphone (eCapture) */
    {
        IMMDevice *device = NULL;
        hr = IMMDeviceEnumerator_GetDefaultAudioEndpoint(enumerator, eCapture, eConsole, &device);
        if (SUCCEEDED(hr) && device) {
            IAudioEndpointVolume *vol = NULL;
            hr = IMMDevice_Activate(device, &IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (void **)&vol);
            if (SUCCEEDED(hr) && vol) {
                float level = 0.0f;
                BOOL muted = FALSE;
                IAudioEndpointVolume_GetMasterVolumeLevelScalar(vol, &level);
                IAudioEndpointVolume_GetMute(vol, &muted);
                g_mic_volume = level;
                g_mic_muted = muted ? 1 : 0;
                IAudioEndpointVolume_Release(vol);
            }
            IMMDevice_Release(device);
        } else {
            g_mic_volume = -1.0f;
            g_mic_muted = 0;
        }
    }

    /* Query speakers (eRender) */
    {
        IMMDevice *device = NULL;
        hr = IMMDeviceEnumerator_GetDefaultAudioEndpoint(enumerator, eRender, eConsole, &device);
        if (SUCCEEDED(hr) && device) {
            IAudioEndpointVolume *vol = NULL;
            hr = IMMDevice_Activate(device, &IID_IAudioEndpointVolume, CLSCTX_ALL, NULL, (void **)&vol);
            if (SUCCEEDED(hr) && vol) {
                float level = 0.0f;
                BOOL muted = FALSE;
                IAudioEndpointVolume_GetMasterVolumeLevelScalar(vol, &level);
                IAudioEndpointVolume_GetMute(vol, &muted);
                g_spk_volume = level;
                g_spk_muted = muted ? 1 : 0;
                IAudioEndpointVolume_Release(vol);
            }
            IMMDevice_Release(device);
        } else {
            g_spk_volume = -1.0f;
            g_spk_muted = 0;
        }
    }

    IMMDeviceEnumerator_Release(enumerator);
}

/* Stats window procedure */
static LRESULT CALLBACK StatsWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);

            RECT rc;
            GetClientRect(hwnd, &rc);
            int h = rc.bottom - rc.top;

            /* Background */
            FillRect(hdc, &rc, g_brush_bg);

            SetBkMode(hdc, TRANSPARENT);

            /* Draw stats - use available height */
            int col_width = (rc.right - rc.left) / 8;
            int label_h = 14;
            int value_top = label_h + 2;

            char buf[32];

            if (g_is_recording) {
                /* Time */
                RECT r1 = { 0, 2, col_width, label_h + 2 };
                SetTextColor(hdc, COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_normal);
                DrawTextA(hdc, "TIME", -1, &r1, DT_CENTER);

                snprintf(buf, sizeof(buf), "%.1fs", g_audio_seconds);
                RECT r1v = { 0, value_top, col_width, h - 2 };
                SetTextColor(hdc, COLOR_ACCENT);
                SelectObject(hdc, g_font_medium);
                DrawTextA(hdc, buf, -1, &r1v, DT_CENTER);

                /* Energy - show SPEAKING or QUIET */
                RECT r2 = { col_width, 2, col_width * 2, label_h + 2 };
                SetTextColor(hdc, COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_normal);
                DrawTextA(hdc, "STATUS", -1, &r2, DT_CENTER);

                int is_speaking = g_current_energy >= SILENCE_THRESHOLD;
                RECT r2v = { col_width, value_top, col_width * 2, h - 2 };
                SetTextColor(hdc, is_speaking ? COLOR_WAVE_LOW : COLOR_SILENCE);
                SelectObject(hdc, g_font_medium);
                DrawTextA(hdc, is_speaking ? "SPEECH" : "QUIET", -1, &r2v, DT_CENTER);

                /* Silence */
                RECT r3 = { col_width * 2, 2, col_width * 3, label_h + 2 };
                SetTextColor(hdc, COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_normal);
                DrawTextA(hdc, "SILENCE", -1, &r3, DT_CENTER);

                snprintf(buf, sizeof(buf), "%d/%d", g_silence_count, SILENCE_CHUNKS);
                RECT r3v = { col_width * 2, value_top, col_width * 3, h - 2 };
                SetTextColor(hdc, g_silence_count > 0 ? COLOR_SILENCE : COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_medium);
                DrawTextA(hdc, buf, -1, &r3v, DT_CENTER);
            } else {
                /* When stopped: show total duration and view position */
                RECT r1 = { 0, 2, col_width, label_h + 2 };
                SetTextColor(hdc, COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_normal);
                DrawTextA(hdc, "TOTAL", -1, &r1, DT_CENTER);

                float total_time = bar_to_time(g_stored_bar_count);
                snprintf(buf, sizeof(buf), "%.1fs", total_time);
                RECT r1v = { 0, value_top, col_width, h - 2 };
                SetTextColor(hdc, COLOR_ACCENT);
                SelectObject(hdc, g_font_medium);
                DrawTextA(hdc, buf, -1, &r1v, DT_CENTER);

                /* View position */
                RECT r2 = { col_width, 2, col_width * 2, label_h + 2 };
                SetTextColor(hdc, COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_normal);
                DrawTextA(hdc, "VIEW", -1, &r2, DT_CENTER);

                float view_start = bar_to_time(g_scroll_offset);
                float view_end = bar_to_time(g_scroll_offset + WAVEFORM_BARS);
                if (view_end > total_time) view_end = total_time;
                snprintf(buf, sizeof(buf), "%.1f-%.1f", view_start, view_end);
                RECT r2v = { col_width, value_top, col_width * 2, h - 2 };
                SetTextColor(hdc, COLOR_TEXT);
                SelectObject(hdc, g_font_medium);
                DrawTextA(hdc, buf, -1, &r2v, DT_CENTER);

                /* Marker */
                RECT r3 = { col_width * 2, 2, col_width * 3, label_h + 2 };
                SetTextColor(hdc, COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_normal);
                DrawTextA(hdc, "MARKER", -1, &r3, DT_CENTER);

                if (g_marker_time >= 0) {
                    snprintf(buf, sizeof(buf), "%.1fs", g_marker_time);
                } else {
                    snprintf(buf, sizeof(buf), "--");
                }
                RECT r3v = { col_width * 2, value_top, col_width * 3, h - 2 };
                SetTextColor(hdc, g_marker_time >= 0 ? RGB(255, 100, 100) : COLOR_TEXT_DIM);
                SelectObject(hdc, g_font_medium);
                DrawTextA(hdc, buf, -1, &r3v, DT_CENTER);
            }

            /* Fifth column: connection + TTS status (always visible) */
            RECT r5 = { col_width * 4, 2, col_width * 5, label_h + 2 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            {
                const char *tts_label;
                if (g_tutor_mode) {
                    tts_label = g_tts_enabled ? "ZH:TTS" : "ZH:tts";
                } else if (g_llm_mode == LLM_MODE_LOCAL) {
                    tts_label = g_tts_enabled ? "LLM/TTS" : "LLM/tts";
                } else {
                    tts_label = g_tts_enabled ? "PIPE/TTS" : "PIPE/tts";
                }
                DrawTextA(hdc, tts_label, -1, &r5, DT_CENTER);
            }

            const char *conn_status;
            COLORREF conn_color;
            if (g_llm_mode == LLM_MODE_LOCAL) {
                if (g_llm_server_ok) {
                    conn_status = "OK";
                    conn_color = COLOR_WAVE_LOW;
                } else {
                    conn_status = "---";
                    conn_color = COLOR_TEXT_DIM;
                }
            } else {
                if (g_pipe_connected) {
                    conn_status = "CONN";
                    conn_color = COLOR_WAVE_LOW;
                } else if (g_pipe_running) {
                    conn_status = "WAIT";
                    conn_color = COLOR_ACCENT;
                } else {
                    conn_status = "---";
                    conn_color = COLOR_TEXT_DIM;
                }
            }
            RECT r5v = { col_width * 4, value_top, col_width * 5, h - 2 };
            SetTextColor(hdc, conn_color);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, conn_status, -1, &r5v, DT_CENTER);

            /* Column 6: MIC volume/mute */
            RECT r6 = { col_width * 5, 2, col_width * 6, label_h + 2 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "MIC", -1, &r6, DT_CENTER);

            const char *mic_text;
            COLORREF mic_color;
            char mic_buf[8];
            if (g_mic_volume < 0.0f) {
                mic_text = "---";
                mic_color = COLOR_TEXT_DIM;
            } else if (g_mic_muted) {
                mic_text = "MUTE";
                mic_color = COLOR_WAVE_HIGH;
            } else {
                snprintf(mic_buf, sizeof(mic_buf), "%d%%", (int)(g_mic_volume * 100.0f));
                mic_text = mic_buf;
                mic_color = COLOR_WAVE_LOW;
            }
            RECT r6v = { col_width * 5, value_top, col_width * 6, h - 2 };
            SetTextColor(hdc, mic_color);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, mic_text, -1, &r6v, DT_CENTER);

            /* Column 7: SPK volume/mute */
            RECT r7 = { col_width * 6, 2, col_width * 7, label_h + 2 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "SPK", -1, &r7, DT_CENTER);

            const char *spk_text;
            COLORREF spk_color;
            char spk_buf[8];
            if (g_spk_volume < 0.0f) {
                spk_text = "---";
                spk_color = COLOR_TEXT_DIM;
            } else if (g_spk_muted) {
                spk_text = "MUTE";
                spk_color = COLOR_WAVE_HIGH;
            } else {
                snprintf(spk_buf, sizeof(spk_buf), "%d%%", (int)(g_spk_volume * 100.0f));
                spk_text = spk_buf;
                spk_color = COLOR_WAVE_LOW;
            }
            RECT r7v = { col_width * 6, value_top, col_width * 7, h - 2 };
            SetTextColor(hdc, spk_color);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, spk_text, -1, &r7v, DT_CENTER);

            /* Column 8: Voice (V to cycle, +/- speed) */
            RECT r8 = { col_width * 7, 2, col_width * 8, label_h + 2 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "VOICE(V)", -1, &r8, DT_CENTER);

            RECT r8v = { col_width * 7, value_top, col_width * 8, h - 2 };
            SelectObject(hdc, g_font_medium);
            {
                int locked_seed = g_tts_voice_seeds[g_tts_voice_idx];
                LONG last_seed = InterlockedCompareExchange(&g_tts_last_seed, 0, 0);
                if (locked_seed >= 0) {
                    /* Locked seed: green */
                    char voice_buf[48];
                    snprintf(voice_buf, sizeof(voice_buf), "%s #%d",
                             g_tts_voices[g_tts_voice_idx], locked_seed);
                    SetTextColor(hdc, COLOR_WAVE_LOW);
                    DrawTextA(hdc, voice_buf, -1, &r8v, DT_CENTER);
                } else if (g_drill_mode && last_seed >= 0) {
                    /* Auditioning: yellow */
                    char voice_buf[48];
                    snprintf(voice_buf, sizeof(voice_buf), "%s ?%ld",
                             g_tts_voices[g_tts_voice_idx], last_seed);
                    SetTextColor(hdc, COLOR_WAVE_MED);
                    DrawTextA(hdc, voice_buf, -1, &r8v, DT_CENTER);
                } else {
                    /* No seed: blue */
                    SetTextColor(hdc, COLOR_ACCENT);
                    DrawTextA(hdc, g_tts_voices[g_tts_voice_idx], -1, &r8v, DT_CENTER);
                }
            }

            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_ERASEBKGND:
            return 1;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* Diagnostics strip window procedure (full-width, below stats row) */
static LRESULT CALLBACK DiagWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);

            RECT rc;
            GetClientRect(hwnd, &rc);
            int h = rc.bottom - rc.top;

            FillRect(hdc, &rc, g_brush_bg);
            SetBkMode(hdc, TRANSPARENT);

            if (!g_is_recording || g_pass_count == 0) {
                EndPaint(hwnd, &ps);
                return 0;
            }

            int col_width = (rc.right - rc.left) / 8;
            int label_h = 14;
            int value_top = label_h + 2;
            char buf[32];

            /* Col 1: PASS */
            RECT s1 = { 0, 1, col_width, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "PASS", -1, &s1, DT_CENTER);

            snprintf(buf, sizeof(buf), "%d", g_pass_count);
            RECT s1v = { 0, value_top, col_width, h - 1 };
            SetTextColor(hdc, g_transcribing ? COLOR_WAVE_MED : COLOR_WAVE_LOW);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s1v, DT_CENTER);

            /* Col 2: RTF */
            RECT s2 = { col_width, 1, col_width * 2, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "RTF", -1, &s2, DT_CENTER);

            snprintf(buf, sizeof(buf), "%.2fx", g_last_rtf);
            RECT s2v = { col_width, value_top, col_width * 2, h - 1 };
            SetTextColor(hdc, g_last_rtf < 1.0 ? COLOR_WAVE_LOW : COLOR_WAVE_HIGH);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s2v, DT_CENTER);

            /* Col 3: WINDOW */
            RECT s3 = { col_width * 2, 1, col_width * 3, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "WINDOW", -1, &s3, DT_CENTER);

            snprintf(buf, sizeof(buf), "%.1fs", g_last_audio_window_sec);
            RECT s3v = { col_width * 2, value_top, col_width * 3, h - 1 };
            SetTextColor(hdc, COLOR_ACCENT);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s3v, DT_CENTER);

            /* Col 4: ENC/DEC */
            RECT s4 = { col_width * 3, 1, col_width * 4, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "ENC/DEC", -1, &s4, DT_CENTER);

            snprintf(buf, sizeof(buf), "%d/%d",
                     (int)g_last_encode_ms, (int)g_last_decode_ms);
            RECT s4v = { col_width * 3, value_top, col_width * 4, h - 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s4v, DT_CENTER);

            /* Col 5: COMMON */
            RECT s5 = { col_width * 4, 1, col_width * 5, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "COMMON", -1, &s5, DT_CENTER);

            snprintf(buf, sizeof(buf), "%d%%", g_last_common_pct);
            RECT s5v = { col_width * 4, value_top, col_width * 5, h - 1 };
            {
                COLORREF common_color;
                if (g_last_common_pct > 60) common_color = COLOR_WAVE_LOW;
                else if (g_last_common_pct >= 30) common_color = COLOR_WAVE_MED;
                else common_color = COLOR_WAVE_HIGH;
                SetTextColor(hdc, common_color);
            }
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s5v, DT_CENTER);

            /* Col 6: COMMIT */
            RECT s6 = { col_width * 5, 1, col_width * 6, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "COMMIT", -1, &s6, DT_CENTER);

            snprintf(buf, sizeof(buf), "%d", g_committed_chars);
            RECT s6v = { col_width * 5, value_top, col_width * 6, h - 1 };
            SetTextColor(hdc, COLOR_ACCENT);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s6v, DT_CENTER);

            /* Col 7: MEM */
            RECT s7 = { col_width * 6, 1, col_width * 7, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "MEM", -1, &s7, DT_CENTER);

            if (g_working_set_mb >= 1024)
                snprintf(buf, sizeof(buf), "%.1fG",
                         (double)g_working_set_mb / 1024.0);
            else
                snprintf(buf, sizeof(buf), "%dM", (int)g_working_set_mb);
            RECT s7v = { col_width * 6, value_top, col_width * 7, h - 1 };
            SetTextColor(hdc, COLOR_TEXT);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s7v, DT_CENTER);

            /* Col 8: CPU */
            RECT s8 = { col_width * 7, 1, col_width * 8, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "CPU", -1, &s8, DT_CENTER);

            snprintf(buf, sizeof(buf), "%d%%", (int)g_cpu_percent);
            RECT s8v = { col_width * 7, value_top, col_width * 8, h - 1 };
            {
                COLORREF cpu_color;
                if (g_cpu_percent < 50.0) cpu_color = COLOR_WAVE_LOW;
                else if (g_cpu_percent < 80.0) cpu_color = COLOR_WAVE_MED;
                else cpu_color = COLOR_WAVE_HIGH;
                SetTextColor(hdc, cpu_color);
            }
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, buf, -1, &s8v, DT_CENTER);

            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_ERASEBKGND:
            return 1;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* System/hardware info strip window procedure */
static LRESULT CALLBACK SysInfoWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);

            RECT rc;
            GetClientRect(hwnd, &rc);
            int h = rc.bottom - rc.top;

            FillRect(hdc, &rc, g_brush_bg);
            SetBkMode(hdc, TRANSPARENT);

            int col_width = (rc.right - rc.left) / 4;
            int label_h = 14;
            int value_top = label_h + 2;

            /* Col 1: OS */
            RECT c1 = { 0, 1, col_width, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "OS", -1, &c1, DT_CENTER);

            RECT c1v = { 0, value_top, col_width, h - 1 };
            SetTextColor(hdc, COLOR_TEXT);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, g_os_version, -1, &c1v, DT_CENTER);

            /* Col 2: CPU */
            RECT c2 = { col_width, 1, col_width * 2, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "CPU", -1, &c2, DT_CENTER);

            RECT c2v = { col_width, value_top, col_width * 2, h - 1 };
            SetTextColor(hdc, COLOR_TEXT);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, g_cpu_name, -1, &c2v, DT_CENTER);

            /* Col 3: CORES */
            RECT c3 = { col_width * 2, 1, col_width * 3, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "CORES", -1, &c3, DT_CENTER);

            RECT c3v = { col_width * 2, value_top, col_width * 3, h - 1 };
            SetTextColor(hdc, COLOR_ACCENT);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, g_cpu_cores, -1, &c3v, DT_CENTER);

            /* Col 4: RAM */
            RECT c4 = { col_width * 3, 1, col_width * 4, label_h + 1 };
            SetTextColor(hdc, COLOR_TEXT_DIM);
            SelectObject(hdc, g_font_normal);
            DrawTextA(hdc, "RAM", -1, &c4, DT_CENTER);

            RECT c4v = { col_width * 3, value_top, col_width * 4, h - 1 };
            SetTextColor(hdc, COLOR_ACCENT);
            SelectObject(hdc, g_font_medium);
            DrawTextA(hdc, g_ram_total, -1, &c4v, DT_CENTER);

            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_ERASEBKGND:
            return 1;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* Update waveform visualization */
static void update_waveform(void) {
    /* Shift live bars left */
    for (int i = 0; i < WAVEFORM_BARS - 1; i++) {
        g_waveform_levels[i] = g_waveform_levels[i + 1];
    }

    EnterCriticalSection(&g_audio_lock);
    float energy = g_current_energy;
    LeaveCriticalSection(&g_audio_lock);

    /* Scale energy (log-ish for better visualization) */
    float level = energy * 15.0f;
    if (level > 1.0f) level = 1.0f;
    g_waveform_levels[WAVEFORM_BARS - 1] = level;

    /* Store this level in the full recording buffer */
    if (g_stored_bar_count < MAX_STORED_BARS) {
        g_stored_levels[g_stored_bar_count++] = level;
    }

    /* Redraw waveform and processing visualizer */
    InvalidateRect(g_hwnd_waveform, NULL, FALSE);


    /* Redraw stats */
    InvalidateRect(g_hwnd_stats, NULL, FALSE);
}

/* Find stable sentence end */
static int find_stable_sentence_end(void) {
    if (g_history_count < STABILITY_COUNT) return -1;

    const char *first = g_transcript_history[0];
    int common_len = (int)strlen(first);

    for (int h = 1; h < g_history_count; h++) {
        const char *other = g_transcript_history[h];
        int i = 0;
        while (i < common_len && first[i] && other[i] && first[i] == other[i]) {
            i++;
        }
        common_len = i;
    }

    if (common_len <= g_finalized_len) return -1;

    int last_boundary = -1;
    for (int i = g_finalized_len; i < common_len; i++) {
        char c = first[i];
        if (c == '.' || c == '?' || c == '!') {
            if (i + 1 >= common_len || first[i + 1] == ' ' || first[i + 1] == '\0') {
                last_boundary = i + 1;
            }
        }
    }

    return last_boundary;
}

/* Trim whitespace */
static void trim_whitespace(char *str) {
    char *start = str;
    while (*start == ' ' || *start == '\t' || *start == '\n' || *start == '\r') start++;
    if (start != str) {
        memmove(str, start, strlen(start) + 1);
    }
    int len = (int)strlen(str);
    while (len > 0 && (str[len-1] == ' ' || str[len-1] == '\t' || str[len-1] == '\n' || str[len-1] == '\r')) {
        str[--len] = '\0';
    }
}

/* Update stability */
static void update_stability(const char *new_transcript) {
    for (int i = MAX_HISTORY - 1; i > 0; i--) {
        strcpy(g_transcript_history[i], g_transcript_history[i - 1]);
    }
    strcpy(g_transcript_history[0], new_transcript);
    trim_whitespace(g_transcript_history[0]);
    if (g_history_count < MAX_HISTORY) g_history_count++;

    int boundary = find_stable_sentence_end();
    const char *current = g_transcript_history[0];

    if (boundary > g_finalized_len) {
        strncpy(g_finalized_text, current, boundary);
        g_finalized_text[boundary] = '\0';
        g_finalized_len = boundary;
    }
}

/* Audio capture thread */
static DWORD WINAPI capture_thread_proc(LPVOID param) {
    (void)param;
    HRESULT hr;
    IMFSourceReader *reader = NULL;
    IMFMediaSource *source = NULL;
    IMFAttributes *attributes = NULL;
    IMFActivate **devices = NULL;
    UINT32 device_count = 0;

    hr = MFCreateAttributes(&attributes, 1);
    if (FAILED(hr)) return 1;

    hr = IMFAttributes_SetGUID(attributes, &MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE,
                               &MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_AUDCAP_GUID);
    if (FAILED(hr)) {
        IMFAttributes_Release(attributes);
        return 1;
    }

    hr = MFEnumDeviceSources(attributes, &devices, &device_count);
    IMFAttributes_Release(attributes);

    if (FAILED(hr) || device_count == 0) return 1;

    hr = IMFActivate_ActivateObject(devices[0], &IID_IMFMediaSource, (void **)&source);
    for (UINT32 i = 0; i < device_count; i++) {
        IMFActivate_Release(devices[i]);
    }
    CoTaskMemFree(devices);

    if (FAILED(hr)) return 1;

    hr = MFCreateSourceReaderFromMediaSource(source, NULL, &reader);
    IMFMediaSource_Release(source);
    if (FAILED(hr)) return 1;

    IMFMediaType *media_type = NULL;
    hr = MFCreateMediaType(&media_type);
    if (SUCCEEDED(hr)) {
        IMFMediaType_SetGUID(media_type, &MF_MT_MAJOR_TYPE, &MFMediaType_Audio);
        IMFMediaType_SetGUID(media_type, &MF_MT_SUBTYPE, &MFAudioFormat_PCM);
        IMFMediaType_SetUINT32(media_type, &MF_MT_AUDIO_SAMPLES_PER_SECOND, WHISPER_SAMPLE_RATE);
        IMFMediaType_SetUINT32(media_type, &MF_MT_AUDIO_NUM_CHANNELS, 1);
        IMFMediaType_SetUINT32(media_type, &MF_MT_AUDIO_BITS_PER_SAMPLE, 16);
        IMFMediaType_SetUINT32(media_type, &MF_MT_AUDIO_BLOCK_ALIGNMENT, 2);
        IMFMediaType_SetUINT32(media_type, &MF_MT_AUDIO_AVG_BYTES_PER_SECOND, WHISPER_SAMPLE_RATE * 2);
        hr = IMFSourceReader_SetCurrentMediaType(reader, (DWORD)MF_SOURCE_READER_FIRST_AUDIO_STREAM,
                                                  NULL, media_type);
        IMFMediaType_Release(media_type);
    }

    if (FAILED(hr)) {
        IMFSourceReader_Release(reader);
        return 1;
    }

    while (g_capture_running) {
        IMFSample *sample = NULL;
        DWORD stream_index, flags;
        LONGLONG timestamp;

        hr = IMFSourceReader_ReadSample(reader, (DWORD)MF_SOURCE_READER_FIRST_AUDIO_STREAM,
                                        0, &stream_index, &flags, &timestamp, &sample);

        if (FAILED(hr) || (flags & MF_SOURCE_READERF_ENDOFSTREAM)) break;

        if (sample) {
            IMFMediaBuffer *buffer = NULL;
            hr = IMFSample_ConvertToContiguousBuffer(sample, &buffer);
            if (SUCCEEDED(hr)) {
                BYTE *audio_data = NULL;
                DWORD audio_length = 0;
                hr = IMFMediaBuffer_Lock(buffer, &audio_data, NULL, &audio_length);
                if (SUCCEEDED(hr)) {
                    add_audio_samples((int16_t *)audio_data, audio_length / 2);
                    IMFMediaBuffer_Unlock(buffer);
                }
                IMFMediaBuffer_Release(buffer);
            }
            IMFSample_Release(sample);
        }
    }

    IMFSourceReader_Release(reader);
    return 0;
}

/* VAD state tracking */
static int g_vad_speech_started = 0;    /* Have we detected speech in current block? */
static int g_vad_silence_chunks = 0;    /* Consecutive silence chunks after speech */
#define VAD_SILENCE_TO_TRANSCRIBE 2     /* Silence chunks needed to trigger transcription */
#define VAD_MIN_SPEECH_SAMPLES (WHISPER_SAMPLE_RATE * 1)  /* Minimum 1 second of audio */

/* Work item for transcription thread */
typedef struct {
    float *samples;   /* heap copy, worker frees */
    int n_samples;
    int is_final;     /* 1 = recording stopped, 0 = periodic update */
} asr_work_t;

/* Per-token message posted from worker thread to UI thread */
typedef struct {
    char text[256];
    int audio_ms;
    int byte_offset;
} AsrTokenMsg;

/* Streaming token buffer: accumulates tokens for interim display */
static char g_token_buf[16384];
static int g_token_buf_len = 0;
static int g_token_chat_anchor = -1;  /* chat_len before first token */

/* Token callback: runs on worker thread, posts to UI thread */
static void asr_stream_token_cb(const char *piece, int audio_ms,
                                 int byte_offset, void *userdata) {
    (void)userdata;
    AsrTokenMsg *msg = (AsrTokenMsg *)malloc(sizeof(AsrTokenMsg));
    if (!msg) return;
    strncpy(msg->text, piece, sizeof(msg->text) - 1);
    msg->text[sizeof(msg->text) - 1] = '\0';
    msg->audio_ms = audio_ms;
    msg->byte_offset = byte_offset;
    PostMessageA(g_hwnd_main, WM_ASR_TOKEN, 0, (LPARAM)msg);
}

static DWORD WINAPI asr_transcribe_thread(LPVOID param) {
    asr_work_t *work = (asr_work_t *)param;
    int is_final = work->is_final;

    AsrResult *result = asr_transcribe_stream(work->samples, work->n_samples,
                                               g_asr_port, g_asr_language,
                                               g_asr_prompt, is_final,
                                               asr_stream_token_cb, NULL);
    free(work->samples);
    free(work);

    if (!result)
        log_event("ASR", "HTTP request failed (server not running?)");

    PostMessageA(g_hwnd_main, WM_TRANSCRIBE_DONE, (WPARAM)is_final, (LPARAM)result);
    return 0;
}

/* Async retranscription for ASR server with sentence stability + sliding window. */
static volatile int g_want_final = 0;
static int g_chat_len_before_interim = -1;
static int g_last_transcribe_samples = 0;

/* Sentence stability detection */
static char g_prev_result[16384];
static int g_prev_result_len = 0;
static int g_stable_len = 0;
static int g_common0_unconfirmed = 0;

#define RETRANSCRIBE_INTERVAL_SAMPLES (WHISPER_SAMPLE_RATE * 3)
#define RETRANSCRIBE_MIN_SAMPLES      (WHISPER_SAMPLE_RATE * 1)

/* Kick a retranscription from committed audio offset to current.
 * is_final=1 when recording has stopped. */
static void asr_kick_retranscribe(int is_final) {
    if (g_transcribing) {
        if (is_final) g_want_final = 1;
        return;
    }
    g_want_final = 0;

    int total = g_recording_samples;
    int start = g_committed_samples;
    int n_samples = total - start;
    if (n_samples < RETRANSCRIBE_MIN_SAMPLES) {
        if (!is_final) return;
        if (g_prev_result_len > 0) {
            /* Promote last interim as final */
            AsrResult *r = (AsrResult *)calloc(1, sizeof(AsrResult));
            if (r) {
                r->text = _strdup(g_prev_result);
                r->is_final = 1;
                log_event("FINAL", "promoting interim (short tail)");
                PostMessageA(g_hwnd_main, WM_TRANSCRIBE_DONE, (WPARAM)1, (LPARAM)r);
            }
        }
        return;
    }

    float *copy = (float *)malloc(n_samples * sizeof(float));
    if (!copy) return;
    memcpy(copy, g_recording_buffer + start, n_samples * sizeof(float));

    asr_work_t *work = (asr_work_t *)malloc(sizeof(asr_work_t));
    if (!work) { free(copy); return; }
    work->samples = copy;
    work->n_samples = n_samples;
    work->is_final = is_final;

    g_window_samples = n_samples;
    g_last_transcribe_samples = total;
    g_transcribing = 1;
    if (g_transcribe_thread) {
        CloseHandle(g_transcribe_thread);
        g_transcribe_thread = NULL;
    }
    HANDLE ht = CreateThread(NULL, 0, asr_transcribe_thread, work, 0, NULL);
    if (ht) {
        g_transcribe_thread = ht;
    } else {
        free(work->samples);
        free(work);
        g_transcribing = 0;
    }
}

/* (qwen-asr direct path removed -- transcription via HTTP to local-ai-server) */

/* Transcribe the current audio block (called when VAD detects end of speech) */
static void transcribe_block(void) {
    static float pcmf32[MAX_AUDIO_SAMPLES];

    int n_samples = get_audio_samples(pcmf32, MAX_AUDIO_SAMPLES);
    if (n_samples < VAD_MIN_SPEECH_SAMPLES) {
        log_event("SKIP", "Block too short, skipping transcription");
        return;
    }

    g_audio_seconds = (float)n_samples / WHISPER_SAMPLE_RATE;

    char buf[64];
    snprintf(buf, sizeof(buf), "Transcribing %.1fs of audio", g_audio_seconds);
    log_event("VAD_BLOCK", buf);

    /* Under retranscribe approach, VAD silence doesn't trigger transcription.
     * Periodic timer handles interim updates; stop_recording handles final.
     * Just clear the circular audio buffer so VAD can detect next speech block. */
    clear_audio_buffer();
}

/* Process completed transcription (called from WM_TRANSCRIBE_DONE) */
static void handle_transcribe_result(char *text) {
    if (!text) return;
    int text_len = (int)strlen(text);
    log_event("TRANSCRIPT", text);

    /* Append to finalized text */
    if (text_len > 0) {
        if (g_finalized_len > 0 && g_finalized_text[g_finalized_len - 1] != ' ') {
            if (g_finalized_len < sizeof(g_finalized_text) - 1) {
                g_finalized_text[g_finalized_len++] = ' ';
                g_finalized_text[g_finalized_len] = '\0';
            }
        }
        int space_left = sizeof(g_finalized_text) - g_finalized_len - 1;
        if (text_len > space_left) text_len = space_left;
        strncpy(g_finalized_text + g_finalized_len, text, text_len);
        g_finalized_len += text_len;
        g_finalized_text[g_finalized_len] = '\0';

        log_event("FINALIZED", g_finalized_text);

        if (g_llm_mode == LLM_MODE_LOCAL) {
            llm_send(text);
        } else {
            pipe_send(text);
        }
        chat_append("You", text);
    }
    InvalidateRect(g_hwnd_stats, NULL, FALSE);
}

/* Check VAD state and trigger transcription when speech ends */
static void check_vad_and_transcribe(void) {
    EnterCriticalSection(&g_audio_lock);
    float energy = g_current_energy;
    int n_samples = g_audio_samples;
    LeaveCriticalSection(&g_audio_lock);

    g_audio_seconds = (float)n_samples / WHISPER_SAMPLE_RATE;

    char buf[128];
    int is_speech = energy >= SILENCE_THRESHOLD;

    if (is_speech) {
        /* Speech detected */
        if (!g_vad_speech_started) {
            log_event("VAD", "Speech started");
            g_vad_speech_started = 1;
        }
        g_vad_silence_chunks = 0;
        g_had_speech = 1;
    } else {
        /* Silence detected */
        if (g_vad_speech_started) {
            g_vad_silence_chunks++;
            snprintf(buf, sizeof(buf), "Silence %d/%d (energy=%.3f, samples=%d)",
                     g_vad_silence_chunks, VAD_SILENCE_TO_TRANSCRIBE, energy, n_samples);
            log_event("VAD", buf);

            /* Enough silence after speech - transcribe this block */
            if (g_vad_silence_chunks >= VAD_SILENCE_TO_TRANSCRIBE) {
                if (n_samples >= VAD_MIN_SPEECH_SAMPLES) {
                    transcribe_block();
                } else {
                    log_event("VAD", "Block too short, discarding");
                    clear_audio_buffer();
                }
                /* Reset for next speech block */
                g_vad_speech_started = 0;
                g_vad_silence_chunks = 0;
            }
        }
    }

    /* Auto-stop after extended silence (no speech for a while) */
    if (!g_vad_speech_started && !is_speech) {
        g_silence_count++;
        if (g_silence_count >= SILENCE_CHUNKS * 2 && g_had_speech && !g_ptt_held) {
            g_pending_stop = 1;
        }
    } else {
        g_silence_count = 0;
    }

    /* Update stats display */
    InvalidateRect(g_hwnd_stats, NULL, FALSE);
}

/* Start/stop recording */
static void start_recording(void) {
    /* Stop any TTS playback so it doesn't talk over the user */
    if (g_tts_voice) {
        ISpVoice_Speak(g_tts_voice, L"", SPF_ASYNC | SPF_PURGEBEFORESPEAK, NULL);
    }
    /* Cancel any pending LLM request */
    InterlockedExchange(&g_llm_interrupt, 1);
    clear_audio_buffer();
    g_history_count = 0;
    g_finalized_len = 0;
    g_finalized_text[0] = '\0';
    g_silence_count = 0;
    g_had_speech = 0;
    g_pending_stop = 0;
    g_audio_seconds = 0.0f;
    memset(g_waveform_levels, 0, sizeof(g_waveform_levels));

    /* Reset stored data for new recording */
    g_stored_bar_count = 0;
    g_scroll_offset = 0;
    g_marker_time = -1.0f;
    g_marker_bar = -1;
    g_last_transcribe_samples = 0;
    g_chat_len_before_interim = -1;
    g_prev_result[0] = '\0';
    g_prev_result_len = 0;
    g_stable_len = 0;
    g_common0_unconfirmed = 0;
    g_committed_samples = 0;
    g_window_samples = 0;
    g_want_final = 0;
    /* Wait for any in-flight transcription thread to finish */
    if (g_transcribe_thread) {
        WaitForSingleObject(g_transcribe_thread, 3000);
        CloseHandle(g_transcribe_thread);
        g_transcribe_thread = NULL;
    }
    /* Drain any pending WM_ASR_TOKEN and WM_TRANSCRIBE_DONE */
    {
        MSG drain;
        while (PeekMessageA(&drain, g_hwnd_main, WM_ASR_TOKEN,
                            WM_ASR_TOKEN, PM_REMOVE)) {
            AsrTokenMsg *tok = (AsrTokenMsg *)drain.lParam;
            free(tok);
        }
        while (PeekMessageA(&drain, g_hwnd_main, WM_TRANSCRIBE_DONE,
                            WM_TRANSCRIBE_DONE, PM_REMOVE)) {
            DispatchMessage(&drain);
        }
    }
    g_token_buf[0] = '\0';
    g_token_buf_len = 0;
    g_token_chat_anchor = -1;
    g_drill_stream_len = 0;
    g_drill_state.has_result = 0;
    if (g_hwnd_drill)
        InvalidateRect(g_hwnd_drill, NULL, FALSE);
    /* Clear prompt bias from previous recording */
    g_asr_prompt[0] = '\0';
    /* Reset resource monitoring stats */
    g_pass_count = 0;
    g_last_transcribe_ms = 0;
    g_last_audio_window_sec = 0;
    g_last_rtf = 0;
    g_last_encode_ms = 0;
    g_last_decode_ms = 0;
    g_last_common_pct = 0;
    g_committed_chars = 0;
    g_cpu_percent = 0;
    g_working_set_mb = 0;
    g_cpu_prev_kernel = 0;
    g_cpu_prev_user = 0;
    g_cpu_prev_time = 0;
    g_recording_samples = 0;  /* Clear full recording buffer */
    g_capture_ready = 0;      /* Will be set once first audio samples arrive */
    g_vad_speech_started = 0;
    g_vad_silence_chunks = 0;

    /* Start new log session */
    if (g_log_file) {
        fprintf(g_log_file, "\n=== NEW RECORDING ===\n");
        fflush(g_log_file);
    }

    for (int i = 0; i < MAX_HISTORY; i++) {
        g_transcript_history[i][0] = '\0';
    }

    g_capture_running = 1;
    g_capture_thread = CreateThread(NULL, 0, capture_thread_proc, NULL, 0, NULL);

    if (g_capture_thread) {
        g_is_recording = 1;
        SetWindowTextA(g_hwnd_btn, "Stop");
        SetWindowTextA(g_hwnd_lbl_audio, "Audio Input:");
        SetTimer(g_hwnd_main, ID_TIMER_TRANSCRIBE, 500, NULL);  /* Check VAD every 500ms */
        SetTimer(g_hwnd_main, ID_TIMER_WAVEFORM, WAVEFORM_UPDATE_MS, NULL);
        /* Hide scrollbar during recording */
        ShowWindow(g_hwnd_scrollbar, SW_HIDE);
    }
}

/* Update scrollbar to match stored data */
static void update_scrollbar(void) {
    if (g_hwnd_scrollbar && g_stored_bar_count > WAVEFORM_BARS) {
        SCROLLINFO si = {0};
        si.cbSize = sizeof(si);
        si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin = 0;
        si.nMax = g_stored_bar_count - 1;
        si.nPage = WAVEFORM_BARS;
        si.nPos = g_scroll_offset;
        SetScrollInfo(g_hwnd_scrollbar, SB_CTL, &si, TRUE);
        EnableWindow(g_hwnd_scrollbar, TRUE);
        ShowWindow(g_hwnd_scrollbar, SW_SHOW);
    } else if (g_hwnd_scrollbar) {
        EnableWindow(g_hwnd_scrollbar, FALSE);
        ShowWindow(g_hwnd_scrollbar, SW_HIDE);
    }
}

static void stop_recording(void) {
    KillTimer(g_hwnd_main, ID_TIMER_TRANSCRIBE);
    KillTimer(g_hwnd_main, ID_TIMER_WAVEFORM);

    g_capture_running = 0;
    if (g_capture_thread) {
        WaitForSingleObject(g_capture_thread, 2000);
        CloseHandle(g_capture_thread);
        g_capture_thread = NULL;
    }

    /* Final retranscription of all recorded audio */
    if (g_recording_samples >= RETRANSCRIBE_MIN_SAMPLES) {
        log_event("STOP", "Final retranscription of all audio");
        asr_kick_retranscribe(1);
    }

    /* Save full recording as WAV for offline testing */
    if (g_recording_samples > 0) {
        write_wav(g_recording_buffer, g_recording_samples);
    }

    /* Signal end of recording session to pipe client (not needed in LLM mode) */
    if (g_llm_mode == LLM_MODE_CLAUDE) {
        pipe_send("__DONE__");
    }

    g_is_recording = 0;
    SetWindowTextA(g_hwnd_btn, "Record");
    SetWindowTextA(g_hwnd_lbl_audio, "Recording (click to set marker):");

    /* Set scroll offset to start of recording to show from beginning */
    g_scroll_offset = 0;

    /* Update scrollbar */
    update_scrollbar();

    InvalidateRect(g_hwnd_waveform, NULL, FALSE);

    InvalidateRect(g_hwnd_stats, NULL, FALSE);
}

#include "drill.c"

/* ---- Word slice playback (uses last-WAV cache) ---- */

static HANDLE g_word_slice_thread = NULL;

typedef struct {
    int16_t *pcm;       /* malloc'd PCM slice */
    int      n_samples;
    int      sr;        /* source sample rate (from WAV header) */
    int      offset_ms; /* ms offset into full audio (for karaoke highlight) */
} WordSliceArgs;

/* Self-contained waveOut playback — uses local handles, not the globals,
 * to avoid racing with the main TTS worker thread. */
static DWORD WINAPI word_slice_thread_proc(LPVOID param) {
    WordSliceArgs *args = (WordSliceArgs *)param;

    /* Open our own waveOut device */
    HANDLE done_event = CreateEventA(NULL, FALSE, FALSE, NULL);
    if (!done_event) goto cleanup;

    int base_rate = (args->sr == 24000) ? 48000 : args->sr;
    int device_rate = base_rate;

    WAVEFORMATEX wfx = {0};
    wfx.wFormatTag = WAVE_FORMAT_PCM;
    wfx.nChannels = 1;
    wfx.nSamplesPerSec = (DWORD)device_rate;
    wfx.wBitsPerSample = 16;
    wfx.nBlockAlign = 2;
    wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * 2;

    HWAVEOUT hwo = NULL;
    if (waveOutOpen(&hwo, WAVE_MAPPER, &wfx,
                    (DWORD_PTR)done_event, 0, CALLBACK_EVENT) != MMSYSERR_NOERROR) {
        CloseHandle(done_event);
        goto cleanup;
    }

    /* Upsample 24kHz → 48kHz (same logic as main path) */
    int16_t *upsampled = NULL;
    const int16_t *play_pcm = args->pcm;
    int play_n = args->n_samples;
    if (base_rate == 48000) {
        upsampled = (int16_t *)malloc((size_t)play_n * 2 * sizeof(int16_t));
        if (upsampled) {
            for (int i = 0; i < play_n; i++) {
                upsampled[i * 2]     = play_pcm[i];
                upsampled[i * 2 + 1] = play_pcm[i];
            }
            play_pcm = upsampled;
            play_n *= 2;
        }
    }

    WAVEHDR hdr = {0};
    hdr.lpData = (LPSTR)play_pcm;
    hdr.dwBufferLength = (DWORD)(play_n * sizeof(int16_t));
    waveOutPrepareHeader(hwo, &hdr, sizeof(hdr));

    ResetEvent(done_event);
    InterlockedExchange(&g_tts_playback_ms, (LONG)args->offset_ms);
    waveOutWrite(hwo, &hdr, sizeof(hdr));

    while (!(hdr.dwFlags & WHDR_DONE)) {
        DWORD ret = WaitForSingleObject(done_event, 50);
        (void)ret;
        if (InterlockedCompareExchange(&g_tts_interrupt, 0, 0)) {
            waveOutReset(hwo);
            break;
        }
        /* Update karaoke position (offset by slice start) */
        if (base_rate > 0) {
            MMTIME mmt = {0};
            mmt.wType = TIME_SAMPLES;
            if (waveOutGetPosition(hwo, &mmt, sizeof(mmt)) == MMSYSERR_NOERROR
                && mmt.wType == TIME_SAMPLES) {
                int pos_ms = (int)((double)mmt.u.sample * 1000.0 / base_rate);
                InterlockedExchange(&g_tts_playback_ms,
                    (LONG)(args->offset_ms + pos_ms));
            }
        }
    }

    InterlockedExchange(&g_tts_playback_ms, -1);
    waveOutUnprepareHeader(hwo, &hdr, sizeof(hdr));
    free(upsampled);
    waveOutReset(hwo);
    waveOutClose(hwo);
    CloseHandle(done_event);

cleanup:
    free(args->pcm);
    free(args);
    PostMessageA(g_hwnd_main, WM_TTS_STATUS, 0, 0);
    return 0;
}

/* Play a time slice from last-WAV cache. If not cached, triggers full fetch. */
static void tts_play_word_slice(int start_ms, int end_ms) {
    if (!g_drill_mode) return;
    int cur_idx = g_drill_state.current_idx;

    /* Try to get cached WAV for current sentence+voice */
    char *wav_copy = NULL;
    int wav_copy_len = 0;
    EnterCriticalSection(&g_tts_last_wav_lock);
    if (g_tts_last_wav && g_tts_last_wav_len > 0
        && g_tts_last_wav_sentence == cur_idx
        && g_tts_last_wav_voice == g_tts_voice_idx) {
        wav_copy = (char *)malloc(g_tts_last_wav_len);
        if (wav_copy) {
            memcpy(wav_copy, g_tts_last_wav, g_tts_last_wav_len);
            wav_copy_len = g_tts_last_wav_len;
        }
    }
    LeaveCriticalSection(&g_tts_last_wav_lock);

    /* No cached WAV — trigger full fetch if idle, else ignore */
    if (!wav_copy) {
        if (g_tts_state == 0) {
            DrillSentence *sent = &g_drill_state.sentences[cur_idx];
            if (sent->chinese[0])
                tts_speak_server(sent->chinese, cur_idx,
                                 g_tts_voice_seeds[g_tts_voice_idx]);
        }
        return;
    }

    /* Parse WAV, extract slice, launch playback */
    const int16_t *pcm = NULL;
    int n_samples = 0, sr = 0;
    if (wav_parse_header(wav_copy, wav_copy_len, &pcm, &n_samples, &sr) != 0) {
        free(wav_copy);
        return;
    }

    int start_sample = (int)((double)start_ms / 1000.0 * sr);
    int end_sample   = (int)((double)end_ms   / 1000.0 * sr);
    if (start_sample < 0) start_sample = 0;
    if (end_sample > n_samples) end_sample = n_samples;
    if (start_sample >= end_sample) { free(wav_copy); return; }

    int slice_n = end_sample - start_sample;
    int16_t *slice_pcm = (int16_t *)malloc(slice_n * sizeof(int16_t));
    if (!slice_pcm) { free(wav_copy); return; }
    memcpy(slice_pcm, pcm + start_sample, slice_n * sizeof(int16_t));
    free(wav_copy);

    /* Interrupt previous word-slice playback (if any).
     * Only set interrupt if a slice thread is running — avoid disturbing the
     * main TTS worker's playback loop which checks the same flag. */
    if (g_word_slice_thread) {
        InterlockedExchange(&g_tts_interrupt, 1);
        WaitForSingleObject(g_word_slice_thread, 2000);
        CloseHandle(g_word_slice_thread);
        g_word_slice_thread = NULL;
        InterlockedExchange(&g_tts_interrupt, 0);
    }

    WordSliceArgs *args = (WordSliceArgs *)malloc(sizeof(WordSliceArgs));
    if (!args) { free(slice_pcm); return; }
    args->pcm = slice_pcm;
    args->n_samples = slice_n;
    args->sr = sr;
    args->offset_ms = start_ms;

    PostMessageA(g_hwnd_main, WM_TTS_STATUS, 2, 0); /* speaking */
    g_word_slice_thread = CreateThread(NULL, 0, word_slice_thread_proc, args, 0, NULL);
    if (!g_word_slice_thread) {
        free(slice_pcm);
        free(args);
        PostMessageA(g_hwnd_main, WM_TTS_STATUS, 0, 0);
    }
}

#include "drill_render.c"

/* Layout constants */
#define MARGIN 8
#define BTN_WIDTH 70
#define BTN_HEIGHT 26
#define STATS_HEIGHT 44
#define DIAG_HEIGHT  38   /* full-width transcription diagnostics strip */
#define SYSINFO_HEIGHT 38 /* system/hardware info strip */
#define LABEL_HEIGHT 14
#define SCROLLBAR_HEIGHT 16

static void do_layout(int width, int height) {
    int x = MARGIN;
    int y = MARGIN;
    int content_width = width - MARGIN * 2;

    /* Button */
    MoveWindow(g_hwnd_btn, x, y, BTN_WIDTH, BTN_HEIGHT, TRUE);

    /* Stats - to the right of button */
    int stats_x = x + BTN_WIDTH + MARGIN;
    int stats_width = content_width - BTN_WIDTH - MARGIN;
    MoveWindow(g_hwnd_stats, stats_x, y, stats_width, STATS_HEIGHT, TRUE);

    y += STATS_HEIGHT + MARGIN;

    /* System info strip (full width, below stats) */
    MoveWindow(g_hwnd_sysinfo, x, y, content_width, SYSINFO_HEIGHT, TRUE);
    y += SYSINFO_HEIGHT + MARGIN;

    /* Diagnostics strip (full width, below sysinfo) */
    if (g_hwnd_diag) {
        MoveWindow(g_hwnd_diag, x, y, content_width, DIAG_HEIGHT, TRUE);
        y += DIAG_HEIGHT + MARGIN;
    }

    /* Layout: waveform + scrollbar, claude response (small), then chat log (rest).
     * 3 labels, scrollbar, 1 text area (claude) */
    int small_text_h = 36;  /* Claude: just show latest */
    int fixed_overhead = (LABEL_HEIGHT + 2) * 3 + SCROLLBAR_HEIGHT + small_text_h + MARGIN * 4;
    int remaining = height - y - fixed_overhead;

    /* Split remaining between waveform and chat log */
    int wave_height = remaining * 25 / 100;
    if (wave_height < 40) wave_height = 40;
    int chat_height = remaining - wave_height;
    if (chat_height < 50) chat_height = 50;

    /* Audio input label */
    MoveWindow(g_hwnd_lbl_audio, x, y, 250, LABEL_HEIGHT, TRUE);
    y += LABEL_HEIGHT + 2;

    /* Live waveform (audio input) */
    MoveWindow(g_hwnd_waveform, x, y, content_width, wave_height, TRUE);
    y += wave_height + MARGIN;

    /* Scrollbar (for panning when stopped) */
    MoveWindow(g_hwnd_scrollbar, x, y, content_width, SCROLLBAR_HEIGHT, TRUE);
    y += SCROLLBAR_HEIGHT + MARGIN;

    /* Claude label */
    MoveWindow(g_hwnd_lbl_claude, x, y, 180, LABEL_HEIGHT, TRUE);
    y += LABEL_HEIGHT + 2;

    if (g_drill_mode && g_hwnd_drill) {
        /* Drill mode: hide claude/chat, show drill panel in combined space */
        ShowWindow(g_hwnd_claude_response, SW_HIDE);
        ShowWindow(g_hwnd_lbl_chat, SW_HIDE);
        ShowWindow(g_hwnd_chat, SW_HIDE);
        int drill_height = small_text_h + MARGIN + LABEL_HEIGHT + 2 + chat_height;
        MoveWindow(g_hwnd_drill, x, y, content_width, drill_height, TRUE);
        ShowWindow(g_hwnd_drill, SW_SHOW);
    } else {
        if (g_hwnd_drill)
            ShowWindow(g_hwnd_drill, SW_HIDE);
        ShowWindow(g_hwnd_claude_response, SW_SHOW);
        ShowWindow(g_hwnd_lbl_chat, SW_SHOW);
        ShowWindow(g_hwnd_chat, SW_SHOW);

        /* Claude response text (small, latest response only) */
        MoveWindow(g_hwnd_claude_response, x, y, content_width, small_text_h, TRUE);
        y += small_text_h + MARGIN;

        /* Chat log label */
        MoveWindow(g_hwnd_lbl_chat, x, y, 200, LABEL_HEIGHT, TRUE);
        y += LABEL_HEIGHT + 2;

        /* Chat log (scrolling conversation history) */
        MoveWindow(g_hwnd_chat, x, y, content_width, chat_height, TRUE);
    }
}

/* Main window procedure */
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_SIZE: {
            int width = LOWORD(lParam);
            int height = HIWORD(lParam);
            if (width > 0 && height > 0) {
                do_layout(width, height);
            }
            return 0;
        }

        case WM_COMMAND:
            if (LOWORD(wParam) == ID_BTN_RECORD) {
                if (g_is_recording) {
                    stop_recording();
                } else {
                    start_recording();
                }
            }
            return 0;

        case WM_TIMER:
            if (wParam == ID_TIMER_TRANSCRIBE && g_is_recording) {
                check_vad_and_transcribe();
                if (g_pending_stop) {
                    g_pending_stop = 0;
                    stop_recording();
                }
                /* Periodic retranscription: kick every ~3s of new audio */
                else if (g_recording_samples - g_last_transcribe_samples
                             >= RETRANSCRIBE_INTERVAL_SAMPLES
                         && g_recording_samples >= RETRANSCRIBE_MIN_SAMPLES) {
                    asr_kick_retranscribe(0);
                }
            } else if (wParam == ID_TIMER_WAVEFORM && g_is_recording) {
                update_waveform();
            } else if (wParam == ID_TIMER_DEVSTATUS) {
                query_device_status();
                if (g_is_recording)
                    query_process_resources();
                InvalidateRect(g_hwnd_stats, NULL, FALSE);
                if (g_hwnd_diag)
                    InvalidateRect(g_hwnd_diag, NULL, FALSE);
                /* Refresh drill panel while prefetch is still running */
                if (g_hwnd_drill && g_drill_mode) {
                    LONG pf_done  = InterlockedCompareExchange(&g_tts_prefetch_done, 0, 0);
                    LONG pf_total = InterlockedCompareExchange(&g_tts_prefetch_total, 0, 0);
                    if (pf_total > 0 && pf_done < pf_total)
                        InvalidateRect(g_hwnd_drill, NULL, FALSE);
                }
            }
            else if (wParam == ID_TIMER_DRILL_FLASH) {
                KillTimer(g_hwnd_main, ID_TIMER_DRILL_FLASH);
                if (g_hwnd_drill)
                    InvalidateRect(g_hwnd_drill, NULL, FALSE);
            }
            else if (wParam == ID_TIMER_DRILL_COPY) {
                KillTimer(g_hwnd_main, ID_TIMER_DRILL_COPY);
                g_drill_copy_row = -1;
                if (g_hwnd_drill)
                    InvalidateRect(g_hwnd_drill, NULL, FALSE);
            }
            else if (wParam == ID_TIMER_PLAYBACK) {
                if (g_drill_mode && g_hwnd_drill)
                    InvalidateRect(g_hwnd_drill, NULL, FALSE);
            }
            return 0;

        case WM_HSCROLL:
            if ((HWND)lParam == g_hwnd_scrollbar && !g_is_recording) {
                SCROLLINFO si = {0};
                si.cbSize = sizeof(si);
                si.fMask = SIF_ALL;
                GetScrollInfo(g_hwnd_scrollbar, SB_CTL, &si);

                int new_pos = si.nPos;
                switch (LOWORD(wParam)) {
                    case SB_LINELEFT:      new_pos -= 1; break;
                    case SB_LINERIGHT:     new_pos += 1; break;
                    case SB_PAGELEFT:      new_pos -= si.nPage; break;
                    case SB_PAGERIGHT:     new_pos += si.nPage; break;
                    case SB_THUMBTRACK:    new_pos = HIWORD(wParam); break;
                    case SB_THUMBPOSITION: new_pos = HIWORD(wParam); break;
                }

                /* Clamp */
                int max_offset = g_stored_bar_count - WAVEFORM_BARS;
                if (max_offset < 0) max_offset = 0;
                if (new_pos < 0) new_pos = 0;
                if (new_pos > max_offset) new_pos = max_offset;

                if (new_pos != g_scroll_offset) {
                    g_scroll_offset = new_pos;
                    si.fMask = SIF_POS;
                    si.nPos = g_scroll_offset;
                    SetScrollInfo(g_hwnd_scrollbar, SB_CTL, &si, TRUE);
                    InvalidateRect(g_hwnd_waveform, NULL, FALSE);
                
                }
            }
            return 0;

        case WM_PIPE_RESPONSE: {
            /* Response text from pipe client (allocated by pipe thread) */
            char *response = (char *)lParam;
            if (response) {
                /* Handle pipe commands (prefixed with __) */
                if (strcmp(response, "__FOCUS__") == 0) {
                    /* Windows blocks SetForegroundWindow from background processes.
                     * Attach to the foreground thread's input to get permission. */
                    HWND fg = GetForegroundWindow();
                    DWORD fg_tid = GetWindowThreadProcessId(fg, NULL);
                    DWORD my_tid = GetCurrentThreadId();
                    if (fg_tid != my_tid) AttachThreadInput(fg_tid, my_tid, TRUE);
                    ShowWindow(g_hwnd_main, SW_RESTORE);
                    SetForegroundWindow(g_hwnd_main);
                    BringWindowToTop(g_hwnd_main);
                    if (fg_tid != my_tid) AttachThreadInput(fg_tid, my_tid, FALSE);
                    log_event("PIPE_CMD", "FOCUS");
                    free(response);
                    return 0;
                }
                SetWindowTextUTF8(g_hwnd_claude_response, response);
                log_event("PIPE_RESP", response);
                chat_append("Claude", response);
                tts_speak(response);
                free(response);
            }
            return 0;
        }

        case WM_LLM_RESPONSE: {
            /* Response text from LLM worker (allocated by worker thread) */
            char *response = (char *)lParam;
            if (response) {
                SetWindowTextUTF8(g_hwnd_claude_response, response);
                log_event("LLM_RESP", response);
                chat_append(g_tutor_mode ? "Tutor" : "LLM", response);
                tts_speak(response);
                free(response);
            }
            return 0;
        }

        case WM_ASR_TOKEN: {
            AsrTokenMsg *tok = (AsrTokenMsg *)lParam;
            if (tok) {
                if (g_drill_mode) {
                    /* Accumulate CJK codepoints + timing for progressive drill.
                     * Skip everything below U+2E80 (ASCII, Latin, special tokens
                     * like <|zh|>, English words from early decode steps). */
                    int cps[64];
                    int n = utf8_to_codepoints(tok->text, cps, 64);
                    for (int i = 0; i < n; i++) {
                        if (cps[i] >= 0x2E80 && !is_strip_cp(cps[i])
                            && g_drill_stream_len < DRILL_MAX_TEXT) {
                            g_drill_stream_ms[g_drill_stream_len] = tok->audio_ms;
                            g_drill_stream_cps[g_drill_stream_len++] = cps[i];
                        }
                    }
                    if (g_hwnd_drill)
                        InvalidateRect(g_hwnd_drill, NULL, FALSE);
                } else {
                    int tlen = (int)strlen(tok->text);
                    if (tlen > 0 && g_token_buf_len + tlen < (int)sizeof(g_token_buf) - 1) {
                        /* Record anchor before first token for rollback */
                        if (g_token_chat_anchor < 0)
                            g_token_chat_anchor = g_chat_len;
                        memcpy(g_token_buf + g_token_buf_len, tok->text, tlen);
                        g_token_buf_len += tlen;
                        g_token_buf[g_token_buf_len] = '\0';

                        /* Roll back previous interim line then show updated tokens */
                        if (g_chat_len_before_interim >= 0) {
                            g_chat_len = g_chat_len_before_interim;
                            g_chat_log[g_chat_len] = '\0';
                        }
                        g_chat_len_before_interim = g_token_chat_anchor;
                        chat_append("...", g_token_buf);
                    }
                }
                free(tok);
            }
            return 0;
        }

        case WM_TRANSCRIBE_DONE: {
            AsrResult *tr = (AsrResult *)lParam;
            int is_final = (int)wParam;
            char *result = tr ? tr->text : NULL;
            int result_len = result ? (int)strlen(result) : 0;
            g_transcribing = 0;

            /* Clear streaming token buffer */
            g_token_buf[0] = '\0';
            g_token_buf_len = 0;
            g_token_chat_anchor = -1;

            /* Populate perf metrics from server response */
            if (tr) {
                g_pass_count++;
                g_last_transcribe_ms = tr->perf_total_ms;
                g_last_audio_window_sec = tr->perf_audio_ms / 1000.0;
                g_last_encode_ms = tr->perf_encode_ms;
                g_last_decode_ms = tr->perf_decode_ms;
                if (tr->perf_audio_ms > 0)
                    g_last_rtf = tr->perf_total_ms / tr->perf_audio_ms;
            }
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            if (g_hwnd_diag)
                InvalidateRect(g_hwnd_diag, NULL, FALSE);

            /* Drill mode: bypass normal stability detection */
            if (g_drill_mode && is_final && result_len > 0) {
                g_drill_flash_tick = GetTickCount();
                g_drill_state.has_result = 0;
                if (g_hwnd_drill)
                    InvalidateRect(g_hwnd_drill, NULL, FALSE);
                UpdateWindow(g_hwnd_drill);

                int match = drill_check(&g_drill_state, result);
                drill_record_attempt(&g_drill_state, match);
                log_event("DRILL", match ? "MATCH" : "MISMATCH");
                SetTimer(g_hwnd_main, ID_TIMER_DRILL_FLASH, DRILL_FLASH_MS, NULL);
                asr_free_result(tr);
                return 0;
            }
            if (g_drill_mode) {
                if (!is_final && g_want_final)
                    asr_kick_retranscribe(1);
                asr_free_result(tr);
                return 0;
            }

            /* Remove previous interim line from chat if present */
            if (g_chat_len_before_interim >= 0) {
                g_chat_len = g_chat_len_before_interim;
                g_chat_log[g_chat_len] = '\0';
                g_chat_len_before_interim = -1;
            }

            if (is_final && result_len > 0) {
                /* Final pass: force-commit everything past g_stable_len */
                if (result_len > g_stable_len) {
                    char *tail = result + g_stable_len;
                    int tlen = result_len - g_stable_len;
                    while (tlen > 0 && tail[tlen - 1] == ' ') tlen--;
                    if (tlen > 0) {
                        char commit_buf[8192];
                        if (tlen >= (int)sizeof(commit_buf))
                            tlen = (int)sizeof(commit_buf) - 1;
                        memcpy(commit_buf, tail, tlen);
                        commit_buf[tlen] = '\0';
                        log_event("FINAL", commit_buf);
                        handle_transcribe_result(commit_buf);
                        g_committed_chars += tlen;
                    }
                } else if (result_len > 0 && g_stable_len == 0) {
                    log_event("FINAL", result);
                    handle_transcribe_result(result);
                    g_committed_chars += result_len;
                }
            } else if (result && result[0]) {
                /* Sentence stability detection */
                int common = 0;
                int did_commit = 0;
                {
                    int ia = 0, ib = 0;
                    while (ia < result_len && ib < g_prev_result_len) {
                        char ca = (char)tolower((unsigned char)result[ia]);
                        char cb = (char)tolower((unsigned char)g_prev_result[ib]);
                        if (ca == '-') ca = ' ';
                        if (cb == '-') cb = ' ';
                        if (ca == ' ' && cb == ' ') {
                            common = ia;
                            ia++; ib++;
                            while (ia < result_len && (result[ia] == ' ' || result[ia] == '-')) ia++;
                            while (ib < g_prev_result_len && (g_prev_result[ib] == ' ' || g_prev_result[ib] == '-')) ib++;
                            continue;
                        }
                        if (ca != cb) break;
                        common = ia + 1;
                        ia++; ib++;
                    }

                    /* Sentence-boundary resync */
                    if (common < result_len && common < g_prev_result_len) {
                        int sb_a = -1;
                        for (int i = common; i < result_len - 1; i++) {
                            if ((result[i] == '.' || result[i] == '!'
                                 || result[i] == '?' || result[i] == ':')
                                && result[i + 1] == ' ') {
                                sb_a = i + 2;
                                break;
                            }
                        }
                        int sb_b = -1;
                        if (sb_a >= 0) {
                            for (int i = common; i < g_prev_result_len - 1; i++) {
                                if ((g_prev_result[i] == '.' || g_prev_result[i] == '!'
                                     || g_prev_result[i] == '?' || g_prev_result[i] == ':')
                                    && g_prev_result[i + 1] == ' ') {
                                    sb_b = i + 2;
                                    break;
                                }
                            }
                        }
                        if (sb_a >= 0 && sb_b >= 0
                            && sb_a < result_len && sb_b < g_prev_result_len) {
                            int ja = sb_a, jb = sb_b;
                            int sync_common = sb_a;
                            while (ja < result_len && jb < g_prev_result_len) {
                                char ca2 = (char)tolower((unsigned char)result[ja]);
                                char cb2 = (char)tolower((unsigned char)g_prev_result[jb]);
                                if (ca2 == '-') ca2 = ' ';
                                if (cb2 == '-') cb2 = ' ';
                                if (ca2 == ' ' && cb2 == ' ') {
                                    sync_common = ja;
                                    ja++; jb++;
                                    while (ja < result_len && (result[ja] == ' ' || result[ja] == '-')) ja++;
                                    while (jb < g_prev_result_len && (g_prev_result[jb] == ' ' || g_prev_result[jb] == '-')) jb++;
                                    continue;
                                }
                                if (ca2 != cb2) break;
                                sync_common = ja + 1;
                                ja++; jb++;
                            }
                            if (sync_common - sb_a >= 20 && sync_common > common) {
                                common = sync_common;
                            }
                        }
                    }
                }

                {
                    int max_len = result_len > g_prev_result_len ? result_len : g_prev_result_len;
                    g_last_common_pct = max_len > 0 ? (common * 100 / max_len) : 0;
                }

                {
                    char log_buf[512];
                    snprintf(log_buf, sizeof(log_buf),
                             "common=%d stable=%d prev_len=%d cur_len=%d",
                             common, g_stable_len, g_prev_result_len, result_len);
                    log_event("STABILITY", log_buf);
                    if (common < result_len && common < g_prev_result_len) {
                        char snippet[64];
                        int sn = result_len - common;
                        if (sn > 30) sn = 30;
                        if (sn >= (int)sizeof(snippet)) sn = (int)sizeof(snippet) - 1;
                        memcpy(snippet, result + common, sn);
                        snippet[sn] = '\0';
                        snprintf(log_buf, sizeof(log_buf),
                                 "diverge at %d: cur=\"%s\" prev=\"%.30s\"",
                                 common, snippet,
                                 g_prev_result + common);
                        log_event("STABILITY", log_buf);
                    }
                }

                int new_stable = g_stable_len;
                int best_comma = -1;
                for (int i = common - 1; i > g_stable_len; i--) {
                    int is_strong = (result[i] == '.' || result[i] == '!'
                                    || result[i] == '?' || result[i] == ':');
                    int is_weak = (result[i] == ',' || result[i] == ';');
                    if (is_strong || is_weak) {
                        if (i + 1 < common && result[i + 1] != ' ')
                            continue;
                        int boundary = i + 1;
                        if (boundary < result_len && result[boundary] == ' ')
                            boundary++;

                        if (is_strong) {
                            new_stable = boundary;
                            break;
                        }
                        if (best_comma < 0
                            && boundary - g_stable_len >= 30
                            && common - boundary >= 15) {
                            best_comma = boundary;
                        }
                    }
                }
                if (new_stable == g_stable_len && best_comma > g_stable_len) {
                    new_stable = best_comma;
                }

                {
                    char log_buf2[128];
                    snprintf(log_buf2, sizeof(log_buf2),
                             "new_stable=%d (was %d), common=%d",
                             new_stable, g_stable_len, common);
                    log_event("STABILITY", log_buf2);
                }

                /* Show newly stable sentences as [You] and advance audio window */
                if (new_stable > g_stable_len) {
                    char sentence[8192];
                    int slen = new_stable - g_stable_len;
                    if (slen >= (int)sizeof(sentence))
                        slen = (int)sizeof(sentence) - 1;
                    memcpy(sentence, result + g_stable_len, slen);
                    while (slen > 0 && sentence[slen - 1] == ' ') slen--;
                    sentence[slen] = '\0';
                    if (slen > 0) {
                        handle_transcribe_result(sentence);
                        g_committed_chars += slen;
                        /* Bias next transcription with committed text */
                        strncpy(g_asr_prompt, sentence, sizeof(g_asr_prompt) - 1);
                        g_asr_prompt[sizeof(g_asr_prompt) - 1] = '\0';
                    }

                    /* Advance audio window using timestamps from server */
                    {
                        int advance = 0;
                        if (tr && tr->timestamps && tr->ts_count > 0) {
                            int last_ms = 0;
                            for (int t = 0; t < tr->ts_count; t++) {
                                if (tr->timestamps[t].byte_offset < new_stable)
                                    last_ms = tr->timestamps[t].audio_ms;
                                else
                                    break;
                            }
                            advance = (int)((long long)last_ms * WHISPER_SAMPLE_RATE / 1000);
                            {
                                char log_buf3[128];
                                snprintf(log_buf3, sizeof(log_buf3),
                                         "advance=%d samples (%dms), committed=%d->%d",
                                         advance, last_ms,
                                         g_committed_samples, g_committed_samples + advance);
                                log_event("WINDOW", log_buf3);
                            }
                        } else if (result_len > 0 && g_window_samples > 0) {
                            advance = (int)((long long)g_window_samples
                                         * new_stable / result_len);
                        }
                        g_committed_samples += advance;
                        if (g_committed_samples > g_last_transcribe_samples)
                            g_committed_samples = g_last_transcribe_samples;
                    }

                    g_stable_len = 0;
                    {
                        int tail_len = result_len - new_stable;
                        if (tail_len > 0
                            && tail_len < (int)sizeof(g_prev_result)) {
                            memmove(g_prev_result, result + new_stable,
                                    tail_len);
                            g_prev_result[tail_len] = '\0';
                            g_prev_result_len = tail_len;
                        } else {
                            g_prev_result[0] = '\0';
                            g_prev_result_len = 0;
                        }
                    }
                    did_commit = 1;
                }

                {
                    {
                        int show_from = did_commit ? new_stable : g_stable_len;
                        if (result_len > show_from) {
                            g_chat_len_before_interim = g_chat_len;
                            chat_append("...", result + show_from);
                        }
                    }

                    if (!did_commit) {
                        int is_full_diverge = (common == 0
                                               && g_prev_result_len > 0
                                               && g_stable_len == 0);
                        if (is_full_diverge && !g_common0_unconfirmed) {
                            g_common0_unconfirmed = 1;
                            log_event("STABILITY",
                                      "common=0 unconfirmed, keeping prev");
                        } else {
                            if (common > 0)
                                g_common0_unconfirmed = 0;
                            if (result_len < (int)sizeof(g_prev_result)) {
                                memcpy(g_prev_result, result, result_len + 1);
                            } else {
                                memcpy(g_prev_result, result,
                                       sizeof(g_prev_result) - 1);
                                g_prev_result[sizeof(g_prev_result) - 1] = '\0';
                            }
                            g_prev_result_len = result_len;
                        }
                    }
                }
            }

            if (!is_final && g_want_final) {
                log_event("STABILITY", "kicking deferred final pass");
                asr_kick_retranscribe(1);
            }

            asr_free_result(tr);
            return 0;
        }

        case WM_TTS_STATUS: {
            g_tts_state = (int)wParam;
            /* Playback highlight timer: repaint drill panel at 50ms while speaking */
            if ((int)wParam == 2 && g_drill_mode) {
                SetTimer(g_hwnd_main, ID_TIMER_PLAYBACK, 50, NULL);
            } else {
                KillTimer(g_hwnd_main, ID_TIMER_PLAYBACK);
            }
            /* Update label to show TTS state with correct mode prefix */
            const char *prefix = g_drill_mode ? "Drill"
                               : g_tutor_mode ? "Tutor"
                               : (g_llm_mode == LLM_MODE_LOCAL) ? "LLM" : "Claude";
            const char *label;
            char label_buf[64];
            switch ((int)wParam) {
                case 1:
                    snprintf(label_buf, sizeof(label_buf), "%s: [generating...]", prefix);
                    label = label_buf;
                    break;
                case 2:
                    snprintf(label_buf, sizeof(label_buf), "%s: [speaking...]", prefix);
                    label = label_buf;
                    break;
                case 3:
                    snprintf(label_buf, sizeof(label_buf), "%s: [TTS server unavailable]", prefix);
                    label = label_buf;
                    break;
                default:
                    snprintf(label_buf, sizeof(label_buf), "%s:", prefix);
                    label = label_buf;
                    break;
            }
            SetWindowTextA(g_hwnd_lbl_claude, label);
            if (g_drill_mode && g_hwnd_drill)
                InvalidateRect(g_hwnd_drill, NULL, FALSE);
            return 0;
        }

        case WM_TTS_CACHED: {
            /* Prefetch completed for a sentence — if it's the current one,
             * publish timestamps so word group separators appear immediately. */
            int cached_idx = (int)wParam;
            if (g_drill_mode && cached_idx == g_drill_state.current_idx) {
                tts_publish_cached_timestamps(cached_idx);
                if (g_hwnd_drill)
                    InvalidateRect(g_hwnd_drill, NULL, FALSE);
            }
            return 0;
        }

        case WM_CTLCOLORSTATIC: {
            HDC hdc = (HDC)wParam;
            SetBkColor(hdc, GetSysColor(COLOR_WINDOW));
            return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
        }

        case WM_DESTROY:
            KillTimer(hwnd, ID_TIMER_DEVSTATUS);
            KillTimer(hwnd, ID_TIMER_PLAYBACK);
            if (g_is_recording) {
                stop_recording();
            }
            /* Shut down pipe thread via shutdown event */
            g_pipe_running = 0;
            if (g_pipe_shutdown_event) {
                SetEvent(g_pipe_shutdown_event);
            }
            if (g_pipe_thread) {
                WaitForSingleObject(g_pipe_thread, 2000);
                CloseHandle(g_pipe_thread);
                g_pipe_thread = NULL;
            }
            if (g_pipe != INVALID_HANDLE_VALUE) {
                CloseHandle(g_pipe);
                g_pipe = INVALID_HANDLE_VALUE;
            }
            if (g_pipe_shutdown_event) {
                CloseHandle(g_pipe_shutdown_event);
                g_pipe_shutdown_event = NULL;
            }
            /* Shut down LLM worker */
            llm_worker_stop();
            /* Shut down server TTS worker */
            tts_worker_stop();
            /* Release SAPI TTS */
            if (g_tts_voice) {
                ISpVoice_Release(g_tts_voice);
                g_tts_voice = NULL;
            }
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance;
    (void)lpCmdLine;

    QueryPerformanceFrequency(&g_freq);
    InitializeCriticalSection(&g_audio_lock);

    /* Open log file */
    g_log_file = fopen("voice_test_gui.log", "a");

    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    MFStartup(MF_VERSION, MFSTARTUP_NOSOCKET);

    /* Create fonts */
    g_font_large = CreateFontA(24, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
    g_font_medium = CreateFontA(18, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                                 DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                 CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
    g_font_normal = CreateFontA(13, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                 DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                 CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
    g_font_small = CreateFontA(11, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
    g_font_italic = CreateFontA(13, 0, 0, 0, FW_NORMAL, TRUE, FALSE, FALSE,
                                 DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                 CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, "Segoe UI");
    g_brush_bg = CreateSolidBrush(COLOR_BG);

    /* Large CJK font for drill mode target display */
    g_font_drill_chinese = CreateFontA(32, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
                                        "Microsoft YaHei");

    /* ASR server: no model to load -- transcription via HTTP */
    if (g_tutor_mode)
        g_asr_language = "Chinese";

    /* Parse --asr-port=N from command line */
    {
        const char *cmd = GetCommandLineA();
        const char *port_arg = strstr(cmd, "--asr-port=");
        if (port_arg) {
            int port = atoi(port_arg + 11);
            if (port > 0 && port < 65536)
                g_asr_port = port;
        }
    }

    /* Resolve drill sentence file path (relative to exe directory) */
    {
        char exe_dir[MAX_PATH];
        GetModuleFileNameA(NULL, exe_dir, MAX_PATH);
        char *last = strrchr(exe_dir, '\\');
        if (last) *last = '\0';
        snprintf(g_drill_sentence_path, sizeof(g_drill_sentence_path),
                 "%s\\..\\data\\drill_sentences.txt",
                 exe_dir);

        /* Progress file in APPDATA */
        const char *appdata = getenv("APPDATA");
        if (appdata) {
            snprintf(g_drill_progress_path, sizeof(g_drill_progress_path),
                     "%s\\local-ai-clients\\drill_progress.txt", appdata);
        }
    }

    /* Register window classes */
    WNDCLASSA wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = "VoiceNoteMain";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassA(&wc);

    WNDCLASSA wc_wave = {0};
    wc_wave.lpfnWndProc = WaveformWndProc;
    wc_wave.hInstance = hInstance;
    wc_wave.hbrBackground = g_brush_bg;
    wc_wave.lpszClassName = "WaveformDisplay";
    wc_wave.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassA(&wc_wave);

    WNDCLASSA wc_stats = {0};
    wc_stats.lpfnWndProc = StatsWndProc;
    wc_stats.hInstance = hInstance;
    wc_stats.hbrBackground = g_brush_bg;
    wc_stats.lpszClassName = "StatsDisplay";
    wc_stats.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassA(&wc_stats);

    WNDCLASSA wc_sysinfo = {0};
    wc_sysinfo.lpfnWndProc = SysInfoWndProc;
    wc_sysinfo.hInstance = hInstance;
    wc_sysinfo.hbrBackground = g_brush_bg;
    wc_sysinfo.lpszClassName = "SysInfoDisplay";
    wc_sysinfo.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassA(&wc_sysinfo);

    WNDCLASSA wc_diag = {0};
    wc_diag.lpfnWndProc = DiagWndProc;
    wc_diag.hInstance = hInstance;
    wc_diag.hbrBackground = g_brush_bg;
    wc_diag.lpszClassName = "DiagDisplay";
    wc_diag.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassA(&wc_diag);

    WNDCLASSA wc_drill = {0};
    wc_drill.lpfnWndProc = DrillWndProc;
    wc_drill.hInstance = hInstance;
    wc_drill.hbrBackground = g_brush_bg;
    wc_drill.lpszClassName = "DrillDisplay";
    wc_drill.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassA(&wc_drill);

    /* Create main window */
    g_hwnd_main = CreateWindowA("VoiceNoteMain", "Voice Note",
                                 WS_OVERLAPPEDWINDOW,
                                 CW_USEDEFAULT, CW_USEDEFAULT, 640, 680,
                                 NULL, NULL, hInstance, NULL);

    /* Record button */
    g_hwnd_btn = CreateWindowA("BUTTON", "Record",
                                WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                0, 0, 80, 28,
                                g_hwnd_main, (HMENU)ID_BTN_RECORD, hInstance, NULL);
    SendMessage(g_hwnd_btn, WM_SETFONT, (WPARAM)g_font_normal, TRUE);

    /* Stats display */
    g_hwnd_stats = CreateWindowA("StatsDisplay", "",
                                  WS_VISIBLE | WS_CHILD | WS_BORDER,
                                  0, 0, 100, 50,
                                  g_hwnd_main, NULL, hInstance, NULL);

    /* System/hardware info strip */
    g_hwnd_sysinfo = CreateWindowA("SysInfoDisplay", "",
                                    WS_VISIBLE | WS_CHILD | WS_BORDER,
                                    0, 0, 100, SYSINFO_HEIGHT,
                                    g_hwnd_main, NULL, hInstance, NULL);

    /* Diagnostics strip */
    g_hwnd_diag = CreateWindowA("DiagDisplay", "",
                                 WS_VISIBLE | WS_CHILD | WS_BORDER,
                                 0, 0, 100, DIAG_HEIGHT,
                                 g_hwnd_main, NULL, hInstance, NULL);

    /* Drill mode panel (initially hidden) */
    g_hwnd_drill = CreateWindowA("DrillDisplay", "",
                                  WS_CHILD | WS_BORDER,
                                  0, 0, 100, 100,
                                  g_hwnd_main, NULL, hInstance, NULL);

    /* Audio input label */
    g_hwnd_lbl_audio = CreateWindowA("STATIC", "Audio Input:",
                                      WS_VISIBLE | WS_CHILD,
                                      0, 0, 120, 16,
                                      g_hwnd_main, NULL, hInstance, NULL);
    SendMessage(g_hwnd_lbl_audio, WM_SETFONT, (WPARAM)g_font_normal, TRUE);

    /* Waveform display (live audio input) */
    g_hwnd_waveform = CreateWindowA("WaveformDisplay", "",
                                     WS_VISIBLE | WS_CHILD | WS_BORDER,
                                     0, 0, 100, 80,
                                     g_hwnd_main, NULL, hInstance, NULL);

    /* Scrollbar for timeline navigation */
    g_hwnd_scrollbar = CreateWindowA("SCROLLBAR", "",
                                      WS_CHILD | SBS_HORZ,  /* Not visible initially */
                                      0, 0, 100, 16,
                                      g_hwnd_main, (HMENU)ID_SCROLLBAR, hInstance, NULL);

    /* Claude response label */
    g_hwnd_lbl_claude = CreateWindowA("STATIC", "Claude:",
                                       WS_VISIBLE | WS_CHILD,
                                       0, 0, 80, 16,
                                       g_hwnd_main, (HMENU)ID_LBL_CLAUDE, hInstance, NULL);
    SendMessage(g_hwnd_lbl_claude, WM_SETFONT, (WPARAM)g_font_normal, TRUE);

    /* Claude response text area (Unicode for CJK support) */
    g_hwnd_claude_response = CreateWindowExW(0, L"EDIT", L"",
                                            WS_VISIBLE | WS_CHILD | WS_BORDER | WS_VSCROLL |
                                            ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                                            0, 0, 100, 100,
                                            g_hwnd_main, (HMENU)ID_EDIT_CLAUDE, hInstance, NULL);
    SendMessage(g_hwnd_claude_response, WM_SETFONT, (WPARAM)g_font_normal, TRUE);

    /* Chat log label */
    g_hwnd_lbl_chat = CreateWindowA("STATIC", "Conversation (Space=talk, T=TTS, L=LLM):",
                                     WS_VISIBLE | WS_CHILD,
                                     0, 0, 280, 16,
                                     g_hwnd_main, (HMENU)ID_LBL_CHAT, hInstance, NULL);
    SendMessage(g_hwnd_lbl_chat, WM_SETFONT, (WPARAM)g_font_normal, TRUE);

    /* Chat log text area (Unicode for CJK support) */
    g_hwnd_chat = CreateWindowExW(0, L"EDIT", L"",
                                 WS_VISIBLE | WS_CHILD | WS_BORDER | WS_VSCROLL |
                                 ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                                 0, 0, 100, 100,
                                 g_hwnd_main, (HMENU)ID_EDIT_CHAT, hInstance, NULL);
    SendMessage(g_hwnd_chat, WM_SETFONT, (WPARAM)g_font_normal, TRUE);

    /* Initialize TTS (SAPI) */
    {
        HRESULT hr = CoCreateInstance(
            &CLSID_SpVoice, NULL, CLSCTX_ALL,
            &IID_ISpVoice, (void **)&g_tts_voice);
        if (FAILED(hr)) {
            log_event("TTS", "Failed to create SpVoice");
            g_tts_voice = NULL;
        } else {
            log_event("TTS", "SpVoice initialized");
        }
    }

    /* Initial layout */
    RECT rc;
    GetClientRect(g_hwnd_main, &rc);
    do_layout(rc.right, rc.bottom);

    /* Start LLM worker thread */
    llm_worker_start();

    /* Start TTS worker thread (server-based) */
    tts_worker_start();

    /* Load per-voice locked seeds */
    tts_seeds_load();

    /* Pre-load drill sentence bank and start background TTS+ASR prefetch */
    if (drill_load_bank(&g_drill_state, g_drill_sentence_path) == 0) {
        char bmsg[80];
        snprintf(bmsg, sizeof(bmsg), "Loaded %d sentences", g_drill_state.num_sentences);
        log_event("DRILL", bmsg);
        tts_groupings_init(g_drill_state.num_sentences);
        tts_prefetch_start();
    }

    /* Start named pipe server thread */
    g_pipe_shutdown_event = CreateEventA(NULL, TRUE, FALSE, NULL);  /* manual reset */
    g_pipe_running = 1;
    g_pipe_thread = CreateThread(NULL, 0, pipe_thread_proc, NULL, 0, NULL);
    if (!g_pipe_thread) {
        log_event("PIPE", "Failed to create pipe thread");
        g_pipe_running = 0;
    }

    /* System/hardware info (one-time) */
    query_system_info();

    /* Initial device status query + 1-second polling timer */
    query_device_status();
    SetTimer(g_hwnd_main, ID_TIMER_DEVSTATUS, 1000, NULL);

    ShowWindow(g_hwnd_main, nCmdShow);
    UpdateWindow(g_hwnd_main);

    /* Message loop - intercept keys before child controls consume them */
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        /* Push-to-talk: spacebar held = record, released = stop */
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_SPACE) {
            /* In drill mode, if last answer was correct, advance first */
            if (g_drill_mode && g_drill_state.has_result
                && g_drill_state.last_diff.match && !g_is_recording) {
                drill_advance(&g_drill_state);
                tts_prefetch_prioritize(g_drill_state.current_idx);
                tts_publish_cached_timestamps(g_drill_state.current_idx);
                g_drill_stream_len = 0;
                if (g_hwnd_drill)
                    InvalidateRect(g_hwnd_drill, NULL, FALSE);
            }
            if (!g_ptt_held && !g_is_recording) {
                /* Interrupt any TTS playback before recording */
                InterlockedExchange(&g_tts_interrupt, 1);
                g_ptt_held = 1;
                g_ptt_start_tick = GetTickCount();
                start_recording();
                log_event("PTT", "Spacebar pressed - recording");
            }
            continue;  /* Don't pass to child controls */
        }
        if (msg.message == WM_KEYUP && msg.wParam == VK_SPACE) {
            if (g_ptt_held) {
                DWORD held_ms = GetTickCount() - g_ptt_start_tick;
                if (held_ms < PTT_MIN_HOLD_MS) {
                    /* Too short — ignore release, let auto-stop handle it */
                    log_event("PTT", "Short press - waiting for auto-stop");
                    g_ptt_held = 0;
                } else {
                    g_ptt_held = 0;
                    if (g_is_recording) {
                        stop_recording();
                        log_event("PTT", "Spacebar released - stopped");
                    }
                }
            }
            continue;
        }
        /* T key toggles TTS (ignore auto-repeat) */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'T'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)) {
            g_tts_enabled = !g_tts_enabled;
            log_event("TTS", g_tts_enabled ? "Enabled" : "Disabled");
            if (!g_tts_enabled) {
                if (g_tts_voice) {
                    ISpVoice_Speak(g_tts_voice, L"", SPF_ASYNC | SPF_PURGEBEFORESPEAK, NULL);
                }
            }
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            continue;
        }
        /* V key: cycle TTS voice (Shift+V = backward, ignore auto-repeat) */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'V'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)) {
            if (GetKeyState(VK_SHIFT) & 0x8000) {
                g_tts_voice_idx = (g_tts_voice_idx + (int)TTS_NUM_VOICES - 1) % (int)TTS_NUM_VOICES;
            } else {
                g_tts_voice_idx = (g_tts_voice_idx + 1) % (int)TTS_NUM_VOICES;
            }
            InterlockedExchange(&g_tts_last_seed, -1);
            tts_last_wav_clear();
            log_event("TTS", g_tts_voices[g_tts_voice_idx]);
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            if (g_drill_mode && g_hwnd_drill)
                InvalidateRect(g_hwnd_drill, NULL, FALSE);
            continue;
        }
        /* Shift+L in drill mode: force fresh fetch and play */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'L'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && (GetKeyState(VK_SHIFT) & 0x8000)
            && g_drill_mode && !g_is_recording) {
            log_event("TTS_SRV", "Shift+L -- speaking fresh");
            tts_last_wav_clear();
            DrillSentence *sent = &g_drill_state.sentences[g_drill_state.current_idx];
            if (sent->chinese[0]) {
                tts_speak_server(sent->chinese, g_drill_state.current_idx,
                                g_tts_voice_seeds[g_tts_voice_idx]);
            }
            continue;
        }
        /* L key: in drill mode, speak target sentence via server TTS (ignore auto-repeat).
         * No-op while already fetching/playing — worker handles replay-vs-fetch via cache. */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'L'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && g_drill_mode && !g_is_recording) {
            if (g_tts_state != 0) continue;  /* busy — ignore */
            DrillSentence *sent = &g_drill_state.sentences[g_drill_state.current_idx];
            if (sent->chinese[0]) {
                log_event("TTS_SRV", "L key -- speaking target sentence");
                tts_speak_server(sent->chinese, g_drill_state.current_idx,
                                g_tts_voice_seeds[g_tts_voice_idx]);
            }
            continue;
        }
        /* L key toggles LLM mode (Shift+L = clear history) -- blocked in tutor mode */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'L'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)) {
            if (g_tutor_mode) {
                log_event("LLM", "L key blocked — tutor mode requires local LLM");
                continue;
            }
            if (GetKeyState(VK_SHIFT) & 0x8000) {
                /* Shift+L: clear conversation history */
                llm_history_clear();
                log_event("LLM", "History cleared");
            } else {
                /* L: toggle between Claude and Local LLM mode */
                if (g_llm_mode == LLM_MODE_CLAUDE) {
                    g_llm_mode = LLM_MODE_LOCAL;
                    SetWindowTextA(g_hwnd_lbl_claude, "LLM:");
                    log_event("LLM", "Switched to LOCAL mode");
                } else {
                    g_llm_mode = LLM_MODE_CLAUDE;
                    SetWindowTextA(g_hwnd_lbl_claude, "Claude:");
                    log_event("LLM", "Switched to CLAUDE mode");
                }
            }
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            continue;
        }
        /* M key toggles Mandarin Tutor mode (ignore auto-repeat) */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'M'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && !(GetKeyState(VK_SHIFT) & 0x8000)) {
            if (g_is_recording) {
                log_event("TUTOR", "M key blocked during recording");
                continue;
            }

            g_tutor_mode = !g_tutor_mode;

            if (g_tutor_mode) {
                log_event("TUTOR", "Entering Mandarin Tutor mode");
                g_asr_language = "Chinese";

                g_tutor_model_loaded = 1;
                g_llm_mode = LLM_MODE_LOCAL;
                llm_history_clear();

                SetWindowTextA(g_hwnd_lbl_claude, "Tutor:");
                log_event("TUTOR", "Mandarin Tutor mode active");
            } else {
                log_event("TUTOR", "Exiting Mandarin Tutor mode");
                g_asr_language = NULL;

                g_tutor_model_loaded = 0;
                llm_history_clear();

                SetWindowTextA(g_hwnd_lbl_claude,
                    (g_llm_mode == LLM_MODE_LOCAL) ? "LLM:" : "Claude:");
                log_event("TUTOR", "English mode restored");
            }
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            continue;
        }
        /* Shift+> (period): regenerate with new random seed (voice tuning) */
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_OEM_PERIOD
            && (GetKeyState(VK_SHIFT) & 0x8000)
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && g_drill_mode && !g_is_recording) {
            DrillSentence *sent = &g_drill_state.sentences[g_drill_state.current_idx];
            if (sent->chinese[0]) {
                log_event("TTS_SRV", "Shift+> -- tuning: regenerate with random seed");
                tts_speak_server(sent->chinese, g_drill_state.current_idx, -2);
            }
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            continue;
        }
        /* Shift+< (comma): lock current seed for this voice */
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_OEM_COMMA
            && (GetKeyState(VK_SHIFT) & 0x8000)
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && g_drill_mode) {
            LONG last = InterlockedCompareExchange(&g_tts_last_seed, 0, 0);
            if (last >= 0) {
                g_tts_voice_seeds[g_tts_voice_idx] = (int)last;
                tts_seeds_save();
                tts_last_wav_clear();
                log_event("TTS_SRV", "Shift+< -- locked seed");
            }
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            continue;
        }
        /* Ctrl+Shift+< (comma): unlock seed for this voice */
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_OEM_COMMA
            && (GetKeyState(VK_SHIFT) & 0x8000)
            && (GetKeyState(VK_CONTROL) & 0x8000)
            && g_drill_mode) {
            g_tts_voice_seeds[g_tts_voice_idx] = -1;
            InterlockedExchange(&g_tts_last_seed, -1);
            tts_seeds_save();
            tts_last_wav_clear();
            log_event("TTS_SRV", "Ctrl+Shift+< -- unlocked seed");
            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            continue;
        }
        /* D key toggles Pronunciation Drill mode (ignore auto-repeat) */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'D'
            && !(msg.lParam & (1 << 30))
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && !(GetKeyState(VK_SHIFT) & 0x8000)) {
            if (g_is_recording) {
                log_event("DRILL", "D key blocked during recording");
                continue;
            }

            g_drill_mode = !g_drill_mode;

            if (g_drill_mode) {
                log_event("DRILL", "Entering Pronunciation Drill mode");
                g_asr_language = "Chinese";

                if (g_drill_state.num_sentences <= 0) {
                    log_event("DRILL", "No sentences loaded — cancelling");
                    g_drill_mode = 0;
                    InvalidateRect(g_hwnd_stats, NULL, FALSE);
                    continue;
                }
                drill_init_game(&g_drill_state, g_drill_progress_path);
                srand(42);
                drill_advance(&g_drill_state);
                g_drill_stream_len = 0;
                tts_prefetch_prioritize(g_drill_state.current_idx);
                tts_publish_cached_timestamps(g_drill_state.current_idx);

                SetWindowTextA(g_hwnd_lbl_claude, "Drill:");
                log_event("DRILL", "Pronunciation Drill active");
            } else {
                log_event("DRILL", "Exiting Pronunciation Drill mode");
                drill_shutdown(&g_drill_state, g_drill_progress_path);

                /* Restore language based on tutor mode */
                g_asr_language = g_tutor_mode ? "Chinese" : NULL;

                SetWindowTextA(g_hwnd_lbl_claude,
                    g_tutor_mode ? "Tutor:" :
                    (g_llm_mode == LLM_MODE_LOCAL) ? "LLM:" : "Claude:");
                log_event("DRILL", "Drill mode exited");
            }

            /* Trigger layout change to show/hide drill panel */
            RECT rc_layout;
            GetClientRect(g_hwnd_main, &rc_layout);
            do_layout(rc_layout.right, rc_layout.bottom);

            InvalidateRect(g_hwnd_stats, NULL, FALSE);
            if (g_hwnd_drill)
                InvalidateRect(g_hwnd_drill, NULL, FALSE);
            continue;
        }
        /* H key cycles HSK level filter in drill mode */
        if (msg.message == WM_KEYDOWN && msg.wParam == 'H'
            && !(GetKeyState(VK_CONTROL) & 0x8000)
            && g_drill_mode) {
            g_drill_state.hsk_filter = (g_drill_state.hsk_filter + 1) % 4;  /* 0,1,2,3 -> all,1,2,3 */
            char hmsg[48];
            snprintf(hmsg, sizeof(hmsg), "HSK filter: %s",
                     g_drill_state.hsk_filter == 0 ? "all" :
                     g_drill_state.hsk_filter == 1 ? "HSK 1" :
                     g_drill_state.hsk_filter == 2 ? "HSK 2" : "HSK 3");
            log_event("DRILL", hmsg);
            drill_advance(&g_drill_state);
            tts_prefetch_prioritize(g_drill_state.current_idx);
            tts_publish_cached_timestamps(g_drill_state.current_idx);
            g_drill_stream_len = 0;
            if (g_hwnd_drill)
                InvalidateRect(g_hwnd_drill, NULL, FALSE);
            continue;
        }
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    /* Cleanup */
    if (g_log_file) {
        fclose(g_log_file);
    }
    /* ASR model cleanup not needed -- transcription via HTTP server */
    if (g_word_slice_thread) {
        WaitForSingleObject(g_word_slice_thread, 1000);
        CloseHandle(g_word_slice_thread);
        g_word_slice_thread = NULL;
    }
    tts_prefetch_stop();
    if (g_drill_mode) {
        drill_shutdown(&g_drill_state, g_drill_progress_path);
    }
    tts_groupings_destroy();
    DeleteObject(g_font_large);
    DeleteObject(g_font_medium);
    DeleteObject(g_font_normal);
    DeleteObject(g_font_small);
    DeleteObject(g_font_italic);
    if (g_font_drill_chinese) DeleteObject(g_font_drill_chinese);
    DeleteObject(g_brush_bg);
    DeleteCriticalSection(&g_audio_lock);
    MFShutdown();
    CoUninitialize();

    return (int)msg.wParam;
}
