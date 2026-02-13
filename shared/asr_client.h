/*
 * asr_client.h - Shared HTTP client for local-ai-server ASR
 *
 * Provides synchronous WinHTTP-based transcription for both the voice note GUI
 * and the test harness. Encodes float32 audio to WAV, POSTs multipart/form-data
 * to /v1/audio/transcriptions, and parses verbose_json responses.
 */
#ifndef ASR_CLIENT_H
#define ASR_CLIENT_H

#include <stddef.h>

typedef struct {
    char *text;
    int is_final;
    int ts_count;
    struct { int byte_offset; int audio_ms; } *timestamps;
    double perf_total_ms, perf_audio_ms, perf_encode_ms, perf_decode_ms;
} AsrResult;

/* Encode float32 samples to WAV buffer (16kHz, 16-bit, mono).
 * Returns malloc'd buffer; caller frees. Sets *out_size. */
unsigned char *asr_encode_wav(const float *samples, int n_samples, size_t *out_size);

/* Build multipart/form-data body with WAV file + optional text fields.
 * Returns malloc'd buffer; caller frees. Sets *out_size.
 * out_boundary must be at least 64 bytes. */
unsigned char *asr_build_multipart(const unsigned char *wav, size_t wav_size,
                                   const char *language, const char *prompt,
                                   size_t *out_size, char *out_boundary,
                                   size_t bnd_size);

/* Parse verbose_json response from ASR server.
 * Returns heap-allocated result or NULL. */
AsrResult *asr_parse_response(const char *json, int json_len, int is_final);

/* Free an AsrResult. */
void asr_free_result(AsrResult *r);

/* Synchronous transcribe: encode to WAV, POST to server, parse response.
 * Returns result or NULL on failure. Caller must asr_free_result(). */
AsrResult *asr_transcribe(const float *samples, int n_samples,
                          int port, const char *language, const char *prompt,
                          int is_final);

/* Per-token streaming callback.
 * piece: decoded token text (UTF-8)
 * audio_ms: estimated audio position in milliseconds
 * byte_offset: byte offset into the accumulating text */
typedef void (*asr_token_cb)(const char *piece, int audio_ms,
                              int byte_offset, void *userdata);

/* Streaming transcribe: same as asr_transcribe but uses SSE to deliver
 * per-token callbacks during inference. Returns final AsrResult on completion.
 * token_cb may be NULL (behaves like asr_transcribe with streaming format). */
AsrResult *asr_transcribe_stream(const float *samples, int n_samples,
                                  int port, const char *language,
                                  const char *prompt, int is_final,
                                  asr_token_cb token_cb, void *userdata);

#endif /* ASR_CLIENT_H */
