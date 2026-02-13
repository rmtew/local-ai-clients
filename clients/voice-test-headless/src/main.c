/*
 * voice-test-headless -- Transcription approach comparison test harness
 *
 * Tests different approaches against real recordings via HTTP to local-ai-server:
 *   1. "retranscribe" -- retranscribe growing audio every N seconds
 *   2. "vad" -- VAD-gated segments (original approach)
 *   3. "timestamps" -- verbose_json response with per-token timestamps
 *   4. "sim" -- full GUI simulation (sliding window + stability detection)
 *
 * Build: clients\voice-test-headless\build.bat
 * Usage: voice-test-headless.exe [options] <recording.wav> [...]
 */

#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <math.h>
#include <windows.h>

#include "asr_client.h"

#define SAMPLE_RATE 16000

/* --- WAV reader --- */
static float *read_wav_f32(const char *path, int *out_samples) {
    FILE *f = fopen(path, "rb");
    if (!f) { fprintf(stderr, "Cannot open %s\n", path); return NULL; }

    char riff[4]; fread(riff, 1, 4, f);
    if (memcmp(riff, "RIFF", 4) != 0) { fclose(f); return NULL; }
    fseek(f, 4, SEEK_CUR);
    char wave[4]; fread(wave, 1, 4, f);
    if (memcmp(wave, "WAVE", 4) != 0) { fclose(f); return NULL; }

    int channels = 0, sample_rate = 0, bits_per_sample = 0, audio_format = 0;
    int data_size = 0;
    unsigned char *raw_data = NULL;

    while (!feof(f)) {
        char chunk_id[4];
        unsigned int chunk_size;
        if (fread(chunk_id, 1, 4, f) != 4) break;
        if (fread(&chunk_size, 4, 1, f) != 1) break;
        if (memcmp(chunk_id, "fmt ", 4) == 0) {
            unsigned short fmt, ch, block_align, bps;
            unsigned int sr, byte_rate;
            fread(&fmt, 2, 1, f); audio_format = fmt;
            fread(&ch, 2, 1, f); channels = ch;
            fread(&sr, 4, 1, f); sample_rate = sr;
            fread(&byte_rate, 4, 1, f);
            fread(&block_align, 2, 1, f);
            fread(&bps, 2, 1, f); bits_per_sample = bps;
            if (chunk_size > 16) fseek(f, chunk_size - 16, SEEK_CUR);
        } else if (memcmp(chunk_id, "data", 4) == 0) {
            data_size = chunk_size;
            raw_data = (unsigned char *)malloc(data_size);
            fread(raw_data, 1, data_size, f);
        } else {
            fseek(f, chunk_size, SEEK_CUR);
        }
    }
    fclose(f);
    if (!raw_data || channels == 0) { free(raw_data); return NULL; }

    int native_samples;
    float *native;
    if (audio_format == 3 && bits_per_sample == 32) {
        native_samples = data_size / (4 * channels);
        native = (float *)malloc(native_samples * sizeof(float));
        float *src = (float *)raw_data;
        for (int i = 0; i < native_samples; i++)
            native[i] = src[i * channels];
    } else if (audio_format == 1 && bits_per_sample == 16) {
        native_samples = data_size / (2 * channels);
        native = (float *)malloc(native_samples * sizeof(float));
        short *src = (short *)raw_data;
        for (int i = 0; i < native_samples; i++)
            native[i] = src[i * channels] / 32768.0f;
    } else {
        fprintf(stderr, "Unsupported WAV: fmt=%d bits=%d\n", audio_format, bits_per_sample);
        free(raw_data); return NULL;
    }
    free(raw_data);

    if (sample_rate != SAMPLE_RATE) {
        int out_n = (int)((long long)native_samples * SAMPLE_RATE / sample_rate);
        float *resampled = (float *)malloc(out_n * sizeof(float));
        for (int i = 0; i < out_n; i++) {
            double src_pos = (double)i * sample_rate / SAMPLE_RATE;
            int idx = (int)src_pos;
            double frac = src_pos - idx;
            if (idx + 1 < native_samples)
                resampled[i] = (float)(native[idx] * (1.0 - frac) + native[idx + 1] * frac);
            else
                resampled[i] = native[idx < native_samples ? idx : native_samples - 1];
        }
        free(native);
        native = resampled;
        native_samples = out_n;
    }

    *out_samples = native_samples;
    return native;
}

/* --- Timer --- */
static LARGE_INTEGER perf_freq;
static double now_ms(void) {
    LARGE_INTEGER t;
    QueryPerformanceCounter(&t);
    return (double)t.QuadPart / (double)perf_freq.QuadPart * 1000.0;
}

/* ========================================================================
 * Approach 1: Retranscribe growing audio every N seconds
 *
 * Simulates real-time: the "wall clock" advances as audio plays.
 * Every interval_sec, we transcribe all audio accumulated so far.
 * The transcription happens in "background" (we don't add its duration
 * to the wall clock -- in the real GUI it would be on a worker thread).
 * ======================================================================== */
static void test_retranscribe(int port, const float *wav, int n_samples,
                               float interval_sec) {
    float duration = (float)n_samples / SAMPLE_RATE;
    int interval_samples = (int)(interval_sec * SAMPLE_RATE);

    printf("--- Retranscribe (interval=%.1fs) ---\n\n", interval_sec);

    /* Walk through audio in interval_sec steps (simulating wall clock) */
    int cursor = 0;
    char *prev_text = NULL;
    double total_transcribe_ms = 0;

    while (cursor < n_samples) {
        cursor += interval_samples;
        if (cursor > n_samples) cursor = n_samples;

        float audio_time = (float)cursor / SAMPLE_RATE;

        double t0 = now_ms();
        AsrResult *r = asr_transcribe(wav, cursor, port, NULL, NULL, 0);
        double elapsed = now_ms() - t0;
        total_transcribe_ms += elapsed;

        char *text = r ? r->text : NULL;

        /* Show what changed */
        if (text && text[0]) {
            int changed = !prev_text || strcmp(text, prev_text) != 0;
            if (changed) {
                printf("[audio %5.1fs, transc %4.0fms] [...] %s\n", audio_time, elapsed, text);
            } else {
                printf("[audio %5.1fs, transc %4.0fms] (unchanged)\n", audio_time, elapsed);
            }
            free(prev_text);
            prev_text = _strdup(text);
        } else {
            printf("[audio %5.1fs, transc %4.0fms] (empty)\n", audio_time, elapsed);
        }
        asr_free_result(r);
    }

    printf("\n  Final: %s\n", prev_text ? prev_text : "(empty)");
    printf("  Total transcription time: %.0fms for %.1fs audio (%.1fx overhead)\n",
           total_transcribe_ms, duration, total_transcribe_ms / (duration * 1000));
    printf("\n");
    free(prev_text);
}

/* ========================================================================
 * Approach 2: VAD-gated segments
 * ======================================================================== */
#define SILENCE_THRESHOLD     0.010f
#define VAD_SILENCE_TO_TRANSCRIBE 2
#define VAD_MIN_SPEECH_SAMPLES (SAMPLE_RATE * 1)
#define VAD_CHECK_MS          500

static void test_vad(int port, const float *wav, int n_samples) {
    float duration = (float)n_samples / SAMPLE_RATE;
    int samples_per_tick = SAMPLE_RATE * VAD_CHECK_MS / 1000;

    printf("--- VAD-gated (threshold=%.3f, silence_chunks=%d) ---\n\n",
           SILENCE_THRESHOLD, VAD_SILENCE_TO_TRANSCRIBE);

    float *audio_buf = (float *)calloc(SAMPLE_RATE * 120, sizeof(float));
    int audio_len = 0;
    int vad_speech = 0, vad_silence = 0;
    int pos = 0;
    double total_transcribe_ms = 0;
    int segment_count = 0;

    while (pos < n_samples) {
        int end = pos + samples_per_tick;
        if (end > n_samples) end = n_samples;
        int chunk = end - pos;

        for (int i = 0; i < chunk; i++)
            audio_buf[audio_len++] = wav[pos + i];

        float energy = 0;
        for (int i = 0; i < chunk; i++)
            energy += fabsf(wav[pos + i]);
        energy /= chunk;

        float t = (float)end / SAMPLE_RATE;
        int is_speech = energy >= SILENCE_THRESHOLD;

        if (is_speech) {
            if (!vad_speech) vad_speech = 1;
            vad_silence = 0;
        } else if (vad_speech) {
            vad_silence++;
            if (vad_silence >= VAD_SILENCE_TO_TRANSCRIBE) {
                if (audio_len >= VAD_MIN_SPEECH_SAMPLES) {
                    double t0 = now_ms();
                    AsrResult *r = asr_transcribe(audio_buf, audio_len, port, NULL, NULL, 0);
                    double elapsed = now_ms() - t0;
                    total_transcribe_ms += elapsed;
                    segment_count++;
                    char *text = r ? r->text : NULL;
                    printf("[audio %5.1fs, %5.1fs seg, %4.0fms] [You] %s\n",
                           t, (float)audio_len / SAMPLE_RATE, elapsed,
                           text ? text : "(empty)");
                    asr_free_result(r);
                }
                audio_len = 0;
                vad_speech = 0;
                vad_silence = 0;
            }
        }
        pos = end;
    }

    if (audio_len >= VAD_MIN_SPEECH_SAMPLES) {
        double t0 = now_ms();
        AsrResult *r = asr_transcribe(audio_buf, audio_len, port, NULL, NULL, 0);
        double elapsed = now_ms() - t0;
        total_transcribe_ms += elapsed;
        segment_count++;
        char *text = r ? r->text : NULL;
        printf("[audio %5.1fs, %5.1fs seg, %4.0fms] [You] %s\n",
               duration, (float)audio_len / SAMPLE_RATE, elapsed,
               text ? text : "(empty)");
        asr_free_result(r);
    }

    printf("\n  Segments: %d, Total transcription: %.0fms for %.1fs audio\n\n",
           segment_count, total_transcribe_ms, duration);
    free(audio_buf);
}

/* ========================================================================
 * Approach 3: Timestamp dump (per-token audio timestamps from server)
 * ======================================================================== */
static void test_timestamps(int port, const float *wav, int n_samples) {
    float duration = (float)n_samples / SAMPLE_RATE;

    printf("--- Timestamps (from server verbose_json) ---\n\n");

    double t0 = now_ms();
    AsrResult *r = asr_transcribe(wav, n_samples, port, NULL, NULL, 0);
    double elapsed = now_ms() - t0;

    if (!r || !r->text || !r->text[0]) {
        printf("  (empty transcription)\n\n");
        asr_free_result(r);
        return;
    }

    printf("  Text: %s\n", r->text);
    printf("  Time: %.0fms for %.1fs audio\n\n", elapsed, duration);

    if (r->ts_count == 0) {
        printf("  No timestamps available.\n\n");
        asr_free_result(r);
        return;
    }

    printf("  %d token timestamps:\n", r->ts_count);
    printf("  %6s  %8s\n", "byte", "audio_ms");
    printf("  %6s  %8s\n", "------", "--------");

    int prev_ms = -1;
    int non_monotonic = 0;

    for (int i = 0; i < r->ts_count; i++) {
        int bo = r->timestamps[i].byte_offset;
        int ms = r->timestamps[i].audio_ms;

        /* Extract the token's text: from byte_offset to next token's byte_offset */
        int text_len = (int)strlen(r->text);
        int next_bo = (i + 1 < r->ts_count) ? r->timestamps[i + 1].byte_offset : text_len;
        if (next_bo > text_len) next_bo = text_len;
        int piece_len = next_bo - bo;
        if (piece_len < 0) piece_len = 0;
        if (bo > text_len) bo = text_len;

        char piece[128];
        if (piece_len >= (int)sizeof(piece)) piece_len = (int)sizeof(piece) - 1;
        if (piece_len > 0) memcpy(piece, r->text + bo, piece_len);
        piece[piece_len] = '\0';

        /* Check monotonicity */
        char flag = ' ';
        if (prev_ms >= 0 && ms < prev_ms) {
            flag = '*';  /* non-monotonic */
            non_monotonic++;
        }
        prev_ms = ms;

        printf("  %6d  %7dms %c \"%s\"\n", r->timestamps[i].byte_offset, ms, flag, piece);
    }

    printf("\n  Monotonicity: %s (%d reversals out of %d tokens)\n",
           non_monotonic == 0 ? "PASS" : "FAIL", non_monotonic, r->ts_count);

    /* Check alignment: first token should be near 0, last near audio duration */
    if (r->ts_count > 0) {
        int first_ms = r->timestamps[0].audio_ms;
        int last_ms = r->timestamps[r->ts_count - 1].audio_ms;
        printf("  Range: %dms - %dms (audio duration: %.0fms)\n",
               first_ms, last_ms, duration * 1000.0f);
    }
    printf("\n");
    asr_free_result(r);
}

/* ========================================================================
 * Approach 4: Full GUI simulation (sliding window + stability detection)
 *
 * Replicates the voice-test-gui main.c WM_TRANSCRIBE_DONE logic exactly:
 *   - Retranscribe from committed_samples every interval
 *   - Fuzzy compare with previous result
 *   - Strong/weak sentence boundary detection
 *   - Timestamp-based window advancement
 *   - Final pass on full remaining audio
 * ======================================================================== */
static void test_sim(int port, const float *wav, int n_samples,
                     float interval_sec) {
    float duration = (float)n_samples / SAMPLE_RATE;
    int interval_samples = (int)(interval_sec * SAMPLE_RATE);
    int min_samples = SAMPLE_RATE * 1;  /* RETRANSCRIBE_MIN_SAMPLES */

    printf("--- GUI Simulation (interval=%.1fs) ---\n\n", interval_sec);

    /* State mirrors voice-test-gui main.c */
    char prev_result[16384] = "";
    int prev_result_len = 0;
    int stable_len = 0;
    int common0_unconfirmed = 0;
    int committed_samples = 0;
    int window_samples = 0;
    int pass_num = 0;
    double total_transcribe_ms = 0;
    char prompt[8192] = "";

    /* Simulate recording: audio arrives in real-time, retranscribe every interval.
     * Track simulated wall-clock time so that transcription duration pushes
     * the next kick point forward (matching the live GUI's behavior where
     * audio keeps accumulating while the worker thread runs). */
    int recording_samples = 0;
    int last_transcribe_samples = 0;
    double sim_clock = 0.0;  /* simulated wall time in seconds */

    while (1) {
        double next_kick_sec = (double)(last_transcribe_samples + interval_samples)
                             / SAMPLE_RATE;
        if (next_kick_sec < sim_clock)
            next_kick_sec = sim_clock;  /* transcription took longer than interval */

        recording_samples = (int)(next_kick_sec * SAMPLE_RATE);
        int is_final = 0;
        if (recording_samples >= n_samples) {
            recording_samples = n_samples;
            is_final = 1;
        }

        /* Build the window: from committed_samples to recording_samples */
        int start = committed_samples;
        int ws = recording_samples - start;
        if (ws < min_samples && !is_final) {
            if (recording_samples >= n_samples) break;
            continue;
        }
        if (ws < min_samples) {
            /* Final pass with too little audio: promote prev_result tail */
            float audio_time = (float)recording_samples / SAMPLE_RATE;
            int tail_len = prev_result_len - stable_len;
            if (tail_len > 0) {
                printf("[%5.1fs FINAL  promote] [You] %s\n",
                       audio_time, prev_result + stable_len);
            } else {
                printf("[%5.1fs FINAL  promote] (nothing new)\n", audio_time);
            }
            break;
        }

        window_samples = ws;
        last_transcribe_samples = recording_samples;
        pass_num++;

        double t0 = now_ms();
        AsrResult *ar = asr_transcribe(wav + start, ws, port, NULL,
                                       prompt[0] ? prompt : NULL, 0);
        double elapsed = now_ms() - t0;
        total_transcribe_ms += elapsed;

        /* Advance simulated clock: transcription completed at kick_time + duration */
        sim_clock = (double)recording_samples / SAMPLE_RATE + elapsed / 1000.0;

        char *result = ar ? ar->text : NULL;
        int result_len = result ? (int)strlen(result) : 0;
        float audio_time = (float)recording_samples / SAMPLE_RATE;

        if (is_final) {
            /* Show whatever remains after stable portion as final [You] */
            if (result_len > stable_len) {
                printf("[%5.1fs FINAL %4.0fms] [You] %s\n",
                       audio_time, elapsed, result + stable_len);
            } else if (result_len > 0 && stable_len == 0) {
                printf("[%5.1fs FINAL %4.0fms] [You] %s\n",
                       audio_time, elapsed, result);
            } else {
                printf("[%5.1fs FINAL %4.0fms] (nothing new)\n",
                       audio_time, elapsed);
            }
            asr_free_result(ar);
            break;
        }

        if (!result || !result[0]) {
            printf("[%5.1fs pass#%d %4.0fms] (empty)\n",
                   audio_time, pass_num, elapsed);
            asr_free_result(ar);
            if (recording_samples >= n_samples) break;
            continue;
        }

        /* --- Fuzzy comparison with sentence-boundary resync --- */
        int common = 0;
        int did_commit = 0;
        {
            int ia = 0, ib = 0;
            while (ia < result_len && ib < prev_result_len) {
                char ca = (char)tolower((unsigned char)result[ia]);
                char cb = (char)tolower((unsigned char)prev_result[ib]);
                if (ca == '-') ca = ' ';
                if (cb == '-') cb = ' ';
                if (ca == ' ' && cb == ' ') {
                    common = ia;
                    ia++; ib++;
                    while (ia < result_len && (result[ia] == ' ' || result[ia] == '-')) ia++;
                    while (ib < prev_result_len && (prev_result[ib] == ' ' || prev_result[ib] == '-')) ib++;
                    continue;
                }
                if (ca != cb) break;
                common = ia + 1;
                ia++; ib++;
            }

            /* Sentence-boundary resync */
            if (common < result_len && common < prev_result_len) {
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
                    for (int i = common; i < prev_result_len - 1; i++) {
                        if ((prev_result[i] == '.' || prev_result[i] == '!'
                             || prev_result[i] == '?' || prev_result[i] == ':')
                            && prev_result[i + 1] == ' ') {
                            sb_b = i + 2;
                            break;
                        }
                    }
                }
                if (sb_a >= 0 && sb_b >= 0
                    && sb_a < result_len && sb_b < prev_result_len) {
                    int ja = sb_a, jb = sb_b;
                    int sync_common = sb_a;
                    while (ja < result_len && jb < prev_result_len) {
                        char ca2 = (char)tolower((unsigned char)result[ja]);
                        char cb2 = (char)tolower((unsigned char)prev_result[jb]);
                        if (ca2 == '-') ca2 = ' ';
                        if (cb2 == '-') cb2 = ' ';
                        if (ca2 == ' ' && cb2 == ' ') {
                            sync_common = ja;
                            ja++; jb++;
                            while (ja < result_len && (result[ja] == ' ' || result[ja] == '-')) ja++;
                            while (jb < prev_result_len && (prev_result[jb] == ' ' || prev_result[jb] == '-')) jb++;
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

        /* --- Find commit boundary (mirrors GUI exactly) --- */
        int new_stable = stable_len;
        int best_comma = -1;
        for (int i = common - 1; i > stable_len; i--) {
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
                    && boundary - stable_len >= 30
                    && common - boundary >= 15) {
                    best_comma = boundary;
                }
            }
        }
        if (new_stable == stable_len && best_comma > stable_len) {
            new_stable = best_comma;
        }

        /* --- Commit stable sentences --- */
        if (new_stable > stable_len) {
            char sentence[8192];
            int slen = new_stable - stable_len;
            if (slen >= (int)sizeof(sentence))
                slen = (int)sizeof(sentence) - 1;
            memcpy(sentence, result + stable_len, slen);
            while (slen > 0 && sentence[slen - 1] == ' ') slen--;
            sentence[slen] = '\0';
            if (slen > 0) {
                printf("[%5.1fs pass#%d %4.0fms] [You] %s\n",
                       audio_time, pass_num, elapsed, sentence);
                /* Bias next transcription with committed text */
                strncpy(prompt, sentence, sizeof(prompt) - 1);
                prompt[sizeof(prompt) - 1] = '\0';
            }

            /* Advance audio window using timestamps from server response */
            {
                int advance = 0;
                if (ar && ar->ts_count > 0) {
                    int last_ms = 0;
                    for (int t = 0; t < ar->ts_count; t++) {
                        if (ar->timestamps[t].byte_offset < new_stable)
                            last_ms = ar->timestamps[t].audio_ms;
                        else
                            break;
                    }
                    advance = (int)((long long)last_ms * SAMPLE_RATE / 1000);
                } else if (result_len > 0 && window_samples > 0) {
                    advance = (int)((long long)window_samples
                                 * new_stable / result_len);
                }
                committed_samples += advance;
                if (committed_samples > last_transcribe_samples)
                    committed_samples = last_transcribe_samples;
                printf("           window: committed=%d (+%d samples, %dms)\n",
                       committed_samples, advance,
                       (int)((long long)committed_samples * 1000 / SAMPLE_RATE));
            }

            /* Keep tail of current result as prev for new window */
            stable_len = 0;
            {
                int tail_len = result_len - new_stable;
                if (tail_len > 0 && tail_len < (int)sizeof(prev_result)) {
                    memmove(prev_result, result + new_stable, tail_len);
                    prev_result[tail_len] = '\0';
                    prev_result_len = tail_len;
                } else {
                    prev_result[0] = '\0';
                    prev_result_len = 0;
                }
            }
            did_commit = 1;
        }

        /* Show interim [...] */
        {
            int show_from = did_commit ? new_stable : stable_len;
            if (result_len > show_from) {
                printf("[%5.1fs pass#%d %4.0fms] [...] %s\n",
                       audio_time, pass_num, elapsed, result + show_from);
            }
        }

        /* Save for next comparison */
        if (!did_commit) {
            int is_full_diverge = (common == 0
                                   && prev_result_len > 0
                                   && stable_len == 0);
            if (is_full_diverge && !common0_unconfirmed) {
                common0_unconfirmed = 1;
                printf("           [common=0 unconfirmed, keeping prev]\n");
            } else {
                if (common > 0)
                    common0_unconfirmed = 0;
                int copy_len = result_len;
                if (copy_len >= (int)sizeof(prev_result))
                    copy_len = (int)sizeof(prev_result) - 1;
                memcpy(prev_result, result, copy_len);
                prev_result[copy_len] = '\0';
                prev_result_len = result_len;
            }
        }

        asr_free_result(ar);
        if (recording_samples >= n_samples) break;
    }

    printf("\n  Total transcription: %.0fms for %.1fs audio (%.1fx overhead)\n\n",
           total_transcribe_ms, duration, total_transcribe_ms / (duration * 1000));
}

/* ========================================================================
 * Main
 * ======================================================================== */
int main(int argc, char **argv) {
    if (argc < 2) {
        fprintf(stderr,
            "Usage: %s [options] <recording.wav> [...]\n"
            "Options:\n"
            "  --mode <retranscribe|vad|timestamps|sim|all>  (default: all)\n"
            "  --interval <sec>   Retranscribe interval (default 2.0)\n"
            "  --port <n>         ASR server port (default 8090)\n",
            argv[0]);
        return 1;
    }

    QueryPerformanceFrequency(&perf_freq);

    const char *mode = "all";
    float interval = 2.0f;
    int port = 8090;
    int first_file = 0;

    for (int i = 1; i < argc; i++) {
        if (strcmp(argv[i], "--mode") == 0 && i + 1 < argc) {
            mode = argv[++i];
        } else if (strcmp(argv[i], "--interval") == 0 && i + 1 < argc) {
            interval = (float)atof(argv[++i]);
        } else if (strcmp(argv[i], "--port") == 0 && i + 1 < argc) {
            port = atoi(argv[++i]);
        } else if (argv[i][0] != '-') {
            if (!first_file) first_file = i;
        }
    }
    if (!first_file) {
        fprintf(stderr, "No input files\n");
        return 1;
    }

    int do_retranscribe = strcmp(mode, "retranscribe") == 0 || strcmp(mode, "all") == 0;
    int do_vad = strcmp(mode, "vad") == 0 || strcmp(mode, "all") == 0;
    int do_timestamps = strcmp(mode, "timestamps") == 0;
    int do_sim = strcmp(mode, "sim") == 0;

    for (int i = first_file; i < argc; i++) {
        if (argv[i][0] == '-') continue;

        int n_samples;
        float *wav = read_wav_f32(argv[i], &n_samples);
        if (!wav) { fprintf(stderr, "Failed: %s\n", argv[i]); continue; }

        float dur = (float)n_samples / SAMPLE_RATE;
        printf("\n================================================\n");
        printf("File: %s (%.1fs, port=%d)\n", argv[i], dur, port);
        printf("================================================\n\n");

        if (do_retranscribe) test_retranscribe(port, wav, n_samples, interval);
        if (do_vad)          test_vad(port, wav, n_samples);
        if (do_timestamps)   test_timestamps(port, wav, n_samples);
        if (do_sim)          test_sim(port, wav, n_samples, interval);

        free(wav);
    }

    return 0;
}
