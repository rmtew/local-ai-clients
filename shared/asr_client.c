/*
 * asr_client.c - Shared HTTP client for local-ai-server ASR
 *
 * Shared HTTP client for local-ai-server ASR. Provides WAV encoding, multipart body
 * construction, JSON response parsing, and synchronous WinHTTP transcription.
 */
#define _CRT_SECURE_NO_WARNINGS
#include "asr_client.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>
#include <winhttp.h>

unsigned char *asr_encode_wav(const float *samples, int n_samples, size_t *out_size) {
    int data_bytes = n_samples * 2;  /* 16-bit PCM */
    int file_size = 44 + data_bytes;
    unsigned char *buf = (unsigned char *)malloc(file_size);
    if (!buf) return NULL;

    /* RIFF header */
    memcpy(buf, "RIFF", 4);
    *(int *)(buf + 4)  = file_size - 8;
    memcpy(buf + 8, "WAVE", 4);

    /* fmt chunk */
    memcpy(buf + 12, "fmt ", 4);
    *(int *)(buf + 16) = 16;        /* chunk size */
    *(short *)(buf + 20) = 1;       /* PCM */
    *(short *)(buf + 22) = 1;       /* mono */
    *(int *)(buf + 24) = 16000;     /* sample rate */
    *(int *)(buf + 28) = 32000;     /* byte rate (16000 * 2) */
    *(short *)(buf + 32) = 2;       /* block align */
    *(short *)(buf + 34) = 16;      /* bits per sample */

    /* data chunk */
    memcpy(buf + 36, "data", 4);
    *(int *)(buf + 40) = data_bytes;

    /* Convert float32 [-1,1] to int16 */
    short *pcm = (short *)(buf + 44);
    for (int i = 0; i < n_samples; i++) {
        float s = samples[i];
        if (s > 1.0f) s = 1.0f;
        if (s < -1.0f) s = -1.0f;
        pcm[i] = (short)(s * 32767.0f);
    }

    *out_size = (size_t)file_size;
    return buf;
}

unsigned char *asr_build_multipart(const unsigned char *wav, size_t wav_size,
                                   const char *language, const char *prompt,
                                   size_t *out_size, char *out_boundary,
                                   size_t bnd_size) {
    /* Generate boundary */
    LARGE_INTEGER ticks;
    QueryPerformanceCounter(&ticks);
    snprintf(out_boundary, bnd_size, "----AsrClient%lld", ticks.QuadPart);

    /* Estimate size: boundary overhead + fields + WAV data */
    size_t est = wav_size + 2048;
    unsigned char *body = (unsigned char *)malloc(est);
    if (!body) return NULL;
    size_t pos = 0;

#define MP_APPEND(fmt, ...) pos += snprintf((char *)body + pos, est - pos, fmt, ##__VA_ARGS__)

    /* file field (binary WAV) */
    MP_APPEND("--%s\r\nContent-Disposition: form-data; name=\"file\"; "
              "filename=\"audio.wav\"\r\nContent-Type: audio/wav\r\n\r\n",
              out_boundary);
    memcpy(body + pos, wav, wav_size);
    pos += wav_size;
    MP_APPEND("\r\n");

    /* response_format */
    MP_APPEND("--%s\r\nContent-Disposition: form-data; name=\"response_format\""
              "\r\n\r\nverbose_json\r\n", out_boundary);

    /* language (optional) */
    if (language && language[0]) {
        MP_APPEND("--%s\r\nContent-Disposition: form-data; name=\"language\""
                  "\r\n\r\n%s\r\n", out_boundary, language);
    }

    /* prompt (optional) */
    if (prompt && prompt[0]) {
        MP_APPEND("--%s\r\nContent-Disposition: form-data; name=\"prompt\""
                  "\r\n\r\n%s\r\n", out_boundary, prompt);
    }

    /* Closing boundary */
    MP_APPEND("--%s--\r\n", out_boundary);
#undef MP_APPEND

    *out_size = pos;
    return body;
}

AsrResult *asr_parse_response(const char *json, int json_len, int is_final) {
    (void)json_len;
    AsrResult *r = (AsrResult *)calloc(1, sizeof(AsrResult));
    if (!r) return NULL;
    r->is_final = is_final;

    /* Extract "text":"..." */
    const char *tp = strstr(json, "\"text\"");
    if (tp) {
        tp = strchr(tp + 6, '"');
        if (tp) {
            tp++;  /* skip opening quote */
            const char *end = tp;
            while (*end && !(*end == '"' && *(end - 1) != '\\')) end++;
            int tlen = (int)(end - tp);
            r->text = (char *)malloc(tlen + 1);
            if (r->text) {
                /* Unescape JSON string */
                int j = 0;
                for (int i = 0; i < tlen; i++) {
                    if (tp[i] == '\\' && i + 1 < tlen) {
                        i++;
                        switch (tp[i]) {
                            case '"':  r->text[j++] = '"'; break;
                            case '\\': r->text[j++] = '\\'; break;
                            case 'n':  r->text[j++] = '\n'; break;
                            case 'r':  r->text[j++] = '\r'; break;
                            case 't':  r->text[j++] = '\t'; break;
                            default:   r->text[j++] = tp[i]; break;
                        }
                    } else {
                        r->text[j++] = tp[i];
                    }
                }
                r->text[j] = '\0';
            }
        }
    }

    /* Extract perf fields */
    {
        const char *p;
        if ((p = strstr(json, "\"perf_total_ms\"")) != NULL)
            r->perf_total_ms = atof(strchr(p + 15, ':') + 1);
        if ((p = strstr(json, "\"perf_encode_ms\"")) != NULL)
            r->perf_encode_ms = atof(strchr(p + 16, ':') + 1);
        if ((p = strstr(json, "\"perf_decode_ms\"")) != NULL)
            r->perf_decode_ms = atof(strchr(p + 16, ':') + 1);
        if ((p = strstr(json, "\"duration\"")) != NULL)
            r->perf_audio_ms = atof(strchr(p + 10, ':') + 1) * 1000.0;
    }

    /* Extract words array for timestamps */
    const char *warr = strstr(json, "\"words\"");
    if (warr) {
        warr = strchr(warr, '[');
        if (warr) {
            /* Count objects in array */
            int count = 0;
            const char *scan = warr;
            while (*scan && *scan != ']') {
                if (*scan == '{') count++;
                scan++;
            }

            if (count > 0) {
                r->timestamps = calloc(count, sizeof(r->timestamps[0]));
                if (r->timestamps) {
                    r->ts_count = count;
                    const char *wp = warr + 1;
                    for (int i = 0; i < count; i++) {
                        wp = strchr(wp, '{');
                        if (!wp) break;
                        const char *obj_end = strchr(wp, '}');
                        if (!obj_end) break;

                        /* Parse byte_offset and audio_ms from this object */
                        const char *bp = strstr(wp, "\"byte_offset\"");
                        if (bp && bp < obj_end)
                            r->timestamps[i].byte_offset = atoi(strchr(bp + 13, ':') + 1);
                        const char *ap = strstr(wp, "\"audio_ms\"");
                        if (ap && ap < obj_end)
                            r->timestamps[i].audio_ms = atoi(strchr(ap + 10, ':') + 1);

                        wp = obj_end + 1;
                    }
                }
            }
        }
    }

    return r;
}

/* Build multipart body with custom response_format field. */
static unsigned char *build_multipart_fmt(const unsigned char *wav, size_t wav_size,
                                          const char *language, const char *prompt,
                                          const char *format,
                                          size_t *out_size, char *out_boundary,
                                          size_t bnd_size) {
    LARGE_INTEGER ticks;
    QueryPerformanceCounter(&ticks);
    snprintf(out_boundary, bnd_size, "----AsrClient%lld", ticks.QuadPart);

    size_t est = wav_size + 2048;
    unsigned char *body = (unsigned char *)malloc(est);
    if (!body) return NULL;
    size_t pos = 0;

#define MP2_APPEND(fmt, ...) pos += snprintf((char *)body + pos, est - pos, fmt, ##__VA_ARGS__)

    MP2_APPEND("--%s\r\nContent-Disposition: form-data; name=\"file\"; "
               "filename=\"audio.wav\"\r\nContent-Type: audio/wav\r\n\r\n",
               out_boundary);
    memcpy(body + pos, wav, wav_size);
    pos += wav_size;
    MP2_APPEND("\r\n");

    MP2_APPEND("--%s\r\nContent-Disposition: form-data; name=\"response_format\""
               "\r\n\r\n%s\r\n", out_boundary, format);

    if (language && language[0]) {
        MP2_APPEND("--%s\r\nContent-Disposition: form-data; name=\"language\""
                   "\r\n\r\n%s\r\n", out_boundary, language);
    }
    if (prompt && prompt[0]) {
        MP2_APPEND("--%s\r\nContent-Disposition: form-data; name=\"prompt\""
                   "\r\n\r\n%s\r\n", out_boundary, prompt);
    }

    MP2_APPEND("--%s--\r\n", out_boundary);
#undef MP2_APPEND

    *out_size = pos;
    return body;
}

/* Parse a token event JSON: {"token":"...","audio_ms":N,"byte_offset":N} */
static int sse_parse_token_event(const char *json, int len,
                                  char *token_out, int token_cap,
                                  int *audio_ms, int *byte_offset) {
    (void)len;
    /* Check for "token" key (token event) vs "done" key (done event) */
    const char *tp = strstr(json, "\"token\"");
    if (!tp) return 0;

    /* Extract token string value */
    tp = strchr(tp + 7, '"');
    if (!tp) return 0;
    tp++; /* skip opening quote */
    int j = 0;
    while (*tp && j < token_cap - 1) {
        if (*tp == '"') break;
        if (*tp == '\\' && *(tp + 1)) {
            tp++;
            switch (*tp) {
                case '"':  token_out[j++] = '"'; break;
                case '\\': token_out[j++] = '\\'; break;
                case 'n':  token_out[j++] = '\n'; break;
                case 'r':  token_out[j++] = '\r'; break;
                case 't':  token_out[j++] = '\t'; break;
                default:   token_out[j++] = *tp; break;
            }
        } else {
            token_out[j++] = *tp;
        }
        tp++;
    }
    token_out[j] = '\0';

    *audio_ms = 0;
    *byte_offset = 0;
    const char *p;
    if ((p = strstr(json, "\"audio_ms\"")) != NULL)
        *audio_ms = atoi(strchr(p + 10, ':') + 1);
    if ((p = strstr(json, "\"byte_offset\"")) != NULL)
        *byte_offset = atoi(strchr(p + 13, ':') + 1);

    return 1;
}

void asr_free_result(AsrResult *r) {
    if (!r) return;
    free(r->text);
    free(r->timestamps);
    free(r);
}

AsrResult *asr_transcribe(const float *samples, int n_samples,
                          int port, const char *language, const char *prompt,
                          int is_final) {
    /* Encode samples to WAV */
    size_t wav_size = 0;
    unsigned char *wav = asr_encode_wav(samples, n_samples, &wav_size);
    if (!wav) return NULL;

    /* Build multipart body */
    char boundary[64];
    size_t body_size = 0;
    unsigned char *body = asr_build_multipart(wav, wav_size, language, prompt,
                                              &body_size, boundary, sizeof(boundary));
    free(wav);
    if (!body) return NULL;

    /* HTTP POST to ASR server */
    HINTERNET hSession = WinHttpOpen(L"AsrClient/1.0",
                                      WINHTTP_ACCESS_TYPE_NO_PROXY,
                                      WINHTTP_NO_PROXY_NAME,
                                      WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) {
        free(body);
        return NULL;
    }

    HINTERNET hConnect = WinHttpConnect(hSession, L"localhost",
                                         (INTERNET_PORT)port, 0);
    if (!hConnect) {
        WinHttpCloseHandle(hSession);
        free(body);
        return NULL;
    }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
                                             L"/v1/audio/transcriptions",
                                             NULL, WINHTTP_NO_REFERER,
                                             WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        free(body);
        return NULL;
    }

    /* Set timeouts: 2s connect, 60s receive (test files can be long) */
    WinHttpSetTimeouts(hRequest, 2000, 2000, 60000, 60000);

    /* Build Content-Type header */
    wchar_t ct_header[256];
    _snwprintf(ct_header, 256,
               L"Content-Type: multipart/form-data; boundary=%hs", boundary);

    BOOL ok = WinHttpSendRequest(hRequest, ct_header, (DWORD)-1L,
                                  body, (DWORD)body_size, (DWORD)body_size, 0);
    free(body);

    if (ok) ok = WinHttpReceiveResponse(hRequest, NULL);

    AsrResult *result = NULL;

    if (ok) {
        /* Read response */
        char *resp_buf = (char *)malloc(65536);
        if (resp_buf) {
            DWORD bytes_read = 0;
            DWORD total = 0;
            while (WinHttpReadData(hRequest, resp_buf + total,
                                    65536 - total - 1, &bytes_read)) {
                if (bytes_read == 0) break;
                total += bytes_read;
                if (total >= 65535) break;
            }
            resp_buf[total] = '\0';
            result = asr_parse_response(resp_buf, (int)total, is_final);
            free(resp_buf);
        }
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    return result;
}

AsrResult *asr_transcribe_stream(const float *samples, int n_samples,
                                  int port, const char *language,
                                  const char *prompt, int is_final,
                                  asr_token_cb token_cb, void *userdata) {
    size_t wav_size = 0;
    unsigned char *wav = asr_encode_wav(samples, n_samples, &wav_size);
    if (!wav) return NULL;

    char boundary[64];
    size_t body_size = 0;
    unsigned char *body = build_multipart_fmt(wav, wav_size, language, prompt,
                                              "streaming_verbose_json",
                                              &body_size, boundary,
                                              sizeof(boundary));
    free(wav);
    if (!body) return NULL;

    HINTERNET hSession = WinHttpOpen(L"AsrClient/1.0",
                                      WINHTTP_ACCESS_TYPE_NO_PROXY,
                                      WINHTTP_NO_PROXY_NAME,
                                      WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) { free(body); return NULL; }

    HINTERNET hConnect = WinHttpConnect(hSession, L"localhost",
                                         (INTERNET_PORT)port, 0);
    if (!hConnect) {
        WinHttpCloseHandle(hSession);
        free(body);
        return NULL;
    }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
                                             L"/v1/audio/transcriptions",
                                             NULL, WINHTTP_NO_REFERER,
                                             WINHTTP_DEFAULT_ACCEPT_TYPES, 0);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        free(body);
        return NULL;
    }

    WinHttpSetTimeouts(hRequest, 2000, 2000, 60000, 60000);

    wchar_t ct_header[256];
    _snwprintf(ct_header, 256,
               L"Content-Type: multipart/form-data; boundary=%hs", boundary);

    BOOL ok = WinHttpSendRequest(hRequest, ct_header, (DWORD)-1L,
                                  body, (DWORD)body_size, (DWORD)body_size, 0);
    free(body);

    if (ok) ok = WinHttpReceiveResponse(hRequest, NULL);

    AsrResult *result = NULL;

    if (ok) {
        /* Read SSE stream incrementally */
        char line_buf[4096];
        int line_pos = 0;
        char done_buf[65536];
        int done_len = 0;
        int got_done = 0;

        for (;;) {
            DWORD avail = 0;
            if (!WinHttpQueryDataAvailable(hRequest, &avail)) break;
            if (avail == 0) break;

            char chunk[4096];
            DWORD to_read = avail < sizeof(chunk) ? avail : sizeof(chunk);
            DWORD bytes_read = 0;
            if (!WinHttpReadData(hRequest, chunk, to_read, &bytes_read)) break;
            if (bytes_read == 0) break;

            /* Process bytes: accumulate lines, handle "data: " prefix */
            for (DWORD i = 0; i < bytes_read; i++) {
                char c = chunk[i];
                if (c == '\n') {
                    line_buf[line_pos] = '\0';
                    /* Strip trailing \r */
                    if (line_pos > 0 && line_buf[line_pos - 1] == '\r')
                        line_buf[--line_pos] = '\0';

                    if (line_pos > 6 && memcmp(line_buf, "data: ", 6) == 0) {
                        const char *payload = line_buf + 6;
                        int payload_len = line_pos - 6;

                        if (strstr(payload, "\"done\"")) {
                            /* Done event -- accumulate for final parse */
                            if (payload_len < (int)sizeof(done_buf) - 1) {
                                memcpy(done_buf, payload, payload_len);
                                done_buf[payload_len] = '\0';
                                done_len = payload_len;
                                got_done = 1;
                            }
                        } else if (token_cb) {
                            /* Token event */
                            char token_text[512];
                            int ams = 0, boff = 0;
                            if (sse_parse_token_event(payload, payload_len,
                                                       token_text,
                                                       sizeof(token_text),
                                                       &ams, &boff)) {
                                token_cb(token_text, ams, boff, userdata);
                            }
                        }
                    }
                    line_pos = 0;
                } else {
                    if (line_pos < (int)sizeof(line_buf) - 1)
                        line_buf[line_pos++] = c;
                }
            }
        }

        if (got_done) {
            result = asr_parse_response(done_buf, done_len, is_final);
        }
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    return result;
}
