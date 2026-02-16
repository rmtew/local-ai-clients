#!/usr/bin/env python3
"""
TTS Narrator -- synthesize long text into a single audio file.

Splits input text into sentence-sized chunks, synthesizes each via the
local-ai-server TTS API, and concatenates the resulting audio into one
continuous WAV file.

Usage:
    python tts_narrator.py input.txt -o output.wav
    python tts_narrator.py input.txt -o output.wav --voice cylon --seed 42
    echo "Hello world." | python tts_narrator.py - -o output.wav
    python tts_narrator.py input.txt -o output.wav --gap 0.3 --max-chars 200

Requires: local-ai-server running with --tts-model (default http://localhost:8090)
"""

import argparse
import io
import json
import os
import re
import struct
import sys
import time
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


def split_sentences(text):
    """Split text into sentences on sentence-ending punctuation.

    Keeps the punctuation attached to the sentence. Handles common
    abbreviations (Mr., Mrs., Dr., etc.) to avoid false splits.
    """
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text.strip())

    if not text:
        return []

    # Split on sentence-ending punctuation followed by space or end
    # Negative lookbehind for common abbreviations
    abbrevs = r'(?<!\bMr)(?<!\bMrs)(?<!\bDr)(?<!\bMs)(?<!\bSt)(?<!\bvs)(?<!\betc)(?<!\bInc)(?<!\bLtd)'
    pattern = abbrevs + r'([.!?]+)\s+'

    parts = re.split(pattern, text)

    # Reassemble: parts alternate between text and punctuation
    sentences = []
    i = 0
    while i < len(parts):
        if i + 1 < len(parts) and re.match(r'^[.!?]+$', parts[i + 1]):
            sentences.append(parts[i] + parts[i + 1])
            i += 2
        else:
            if parts[i].strip():
                sentences.append(parts[i])
            i += 1

    return sentences


def merge_short_sentences(sentences, min_chars=40):
    """Merge consecutive short sentences to avoid tiny audio clips."""
    if not sentences:
        return []

    merged = [sentences[0]]
    for s in sentences[1:]:
        if len(merged[-1]) < min_chars:
            merged[-1] = merged[-1] + ' ' + s
        else:
            merged.append(s)
    return merged


def split_long_sentences(sentences, max_chars=300):
    """Split sentences that are too long on clause boundaries."""
    result = []
    for s in sentences:
        if len(s) <= max_chars:
            result.append(s)
            continue

        # Try splitting on comma, semicolon, colon, or dash
        parts = re.split(r'(,\s+|;\s+|:\s+|\s+--\s+)', s)
        current = ''
        for part in parts:
            if len(current) + len(part) > max_chars and current:
                result.append(current.strip())
                current = part
            else:
                current += part
        if current.strip():
            result.append(current.strip())

    return result


def chunk_text(text, min_chars=40, max_chars=300):
    """Split text into synthesis-ready chunks."""
    sentences = split_sentences(text)
    sentences = split_long_sentences(sentences, max_chars)
    sentences = merge_short_sentences(sentences, min_chars)
    return sentences


def synthesize_chunk(server_url, text, voice=None, seed=None,
                     temperature=0.9, speed=1.0):
    """Synthesize a single text chunk. Returns raw WAV bytes."""
    payload = {'input': text}
    if voice:
        payload['voice'] = voice
    if seed is not None:
        payload['seed'] = seed
    if temperature != 0.9:
        payload['temperature'] = temperature
    if speed != 1.0:
        payload['speed'] = speed

    data = json.dumps(payload).encode()
    req = Request(f"{server_url}/v1/audio/speech",
                  data=data,
                  headers={'Content-Type': 'application/json'})

    resp = urlopen(req, timeout=300)
    return resp.read()


def parse_wav(wav_bytes):
    """Parse WAV file. Returns (sample_rate, bits_per_sample, n_channels, pcm_data)."""
    if wav_bytes[:4] != b'RIFF' or wav_bytes[8:12] != b'WAVE':
        raise ValueError("Not a valid WAV file")

    sample_rate = struct.unpack('<I', wav_bytes[24:28])[0]
    n_channels = struct.unpack('<H', wav_bytes[22:24])[0]
    bits = struct.unpack('<H', wav_bytes[34:36])[0]

    # Find data chunk
    idx = wav_bytes.index(b'data')
    data_size = struct.unpack('<I', wav_bytes[idx + 4:idx + 8])[0]
    pcm_data = wav_bytes[idx + 8:idx + 8 + data_size]

    return sample_rate, bits, n_channels, pcm_data


def make_wav_header(sample_rate, bits_per_sample, n_channels, data_size):
    """Create a WAV file header."""
    byte_rate = sample_rate * n_channels * bits_per_sample // 8
    block_align = n_channels * bits_per_sample // 8
    wav_size = 36 + data_size

    header = bytearray()
    header += b'RIFF'
    header += struct.pack('<I', wav_size)
    header += b'WAVE'
    header += b'fmt '
    header += struct.pack('<I', 16)  # chunk size
    header += struct.pack('<H', 1)   # PCM
    header += struct.pack('<H', n_channels)
    header += struct.pack('<I', sample_rate)
    header += struct.pack('<I', byte_rate)
    header += struct.pack('<H', block_align)
    header += struct.pack('<H', bits_per_sample)
    header += b'data'
    header += struct.pack('<I', data_size)
    return bytes(header)


def make_silence(sample_rate, bits_per_sample, n_channels, duration_s):
    """Create silence PCM data for the given duration."""
    n_samples = int(sample_rate * duration_s)
    bytes_per_sample = bits_per_sample // 8
    return b'\x00' * (n_samples * n_channels * bytes_per_sample)


def main():
    parser = argparse.ArgumentParser(
        description='Synthesize long text into a single audio file')
    parser.add_argument('input', help='Input text file (or - for stdin)')
    parser.add_argument('-o', '--output', required=True,
                        help='Output WAV file path')
    parser.add_argument('--server', default='http://localhost:8090',
                        help='Server URL (default: http://localhost:8090)')
    parser.add_argument('--voice', default=None,
                        help='Voice preset name')
    parser.add_argument('--seed', type=int, default=None,
                        help='Random seed for deterministic output')
    parser.add_argument('--temperature', type=float, default=0.9,
                        help='Sampling temperature (default: 0.9)')
    parser.add_argument('--speed', type=float, default=1.0,
                        help='Playback speed multiplier (default: 1.0)')
    parser.add_argument('--gap', type=float, default=0.4,
                        help='Silence gap between sentences in seconds (default: 0.4)')
    parser.add_argument('--min-chars', type=int, default=40,
                        help='Merge sentences shorter than this (default: 40)')
    parser.add_argument('--max-chars', type=int, default=300,
                        help='Split sentences longer than this (default: 300)')
    parser.add_argument('--dry-run', action='store_true',
                        help='Show chunks without synthesizing')
    args = parser.parse_args()

    # Read input
    if args.input == '-':
        text = sys.stdin.read()
    else:
        with open(args.input, encoding='utf-8') as f:
            text = f.read()

    if not text.strip():
        print("Error: empty input")
        sys.exit(1)

    # Split into chunks
    chunks = chunk_text(text, args.min_chars, args.max_chars)

    if not chunks:
        print("Error: no synthesizable text found")
        sys.exit(1)

    total_chars = sum(len(c) for c in chunks)
    print(f"Input: {len(text)} chars -> {len(chunks)} chunks ({total_chars} chars)")

    if args.dry_run:
        for i, chunk in enumerate(chunks):
            print(f"  [{i+1:3d}] ({len(chunk):3d} chars) {chunk[:80]}{'...' if len(chunk) > 80 else ''}")
        sys.exit(0)

    # Verify server
    try:
        urlopen(f"{args.server}/health", timeout=5)
    except Exception:
        print(f"Error: server not reachable at {args.server}")
        sys.exit(1)

    # Synthesize each chunk
    pcm_segments = []
    sample_rate = None
    bits_per_sample = None
    n_channels = None
    total_duration = 0.0
    t_start = time.time()

    for i, chunk in enumerate(chunks):
        label = chunk[:60] + ('...' if len(chunk) > 60 else '')
        print(f"  [{i+1}/{len(chunks)}] \"{label}\"", end='', flush=True)

        t0 = time.time()
        try:
            wav_bytes = synthesize_chunk(
                args.server, chunk,
                voice=args.voice, seed=args.seed,
                temperature=args.temperature, speed=args.speed)
        except HTTPError as e:
            print(f" ERROR {e.code}: {e.read().decode()}")
            continue
        except Exception as e:
            print(f" ERROR: {e}")
            continue

        sr, bits, ch, pcm = parse_wav(wav_bytes)
        elapsed = time.time() - t0
        chunk_duration = len(pcm) / (sr * ch * bits // 8)

        # Verify consistent format
        if sample_rate is None:
            sample_rate = sr
            bits_per_sample = bits
            n_channels = ch
        elif sr != sample_rate or bits != bits_per_sample or ch != n_channels:
            print(f" WARNING: format mismatch ({sr}/{bits}/{ch} vs {sample_rate}/{bits_per_sample}/{n_channels})")
            continue

        pcm_segments.append(pcm)
        total_duration += chunk_duration

        # Add silence gap between sentences (not after last)
        if i < len(chunks) - 1 and args.gap > 0:
            silence = make_silence(sample_rate, bits_per_sample, n_channels, args.gap)
            pcm_segments.append(silence)
            total_duration += args.gap

        print(f" {chunk_duration:.1f}s ({elapsed:.1f}s)")

    if not pcm_segments:
        print("Error: no audio generated")
        sys.exit(1)

    # Concatenate and write output
    all_pcm = b''.join(pcm_segments)
    header = make_wav_header(sample_rate, bits_per_sample, n_channels, len(all_pcm))

    with open(args.output, 'wb') as f:
        f.write(header)
        f.write(all_pcm)

    total_time = time.time() - t_start
    file_size = os.path.getsize(args.output)

    print(f"\nOutput: {args.output}")
    print(f"  Duration: {total_duration:.1f}s ({total_duration/60:.1f} min)")
    print(f"  File size: {file_size / 1024:.0f} KB")
    print(f"  Chunks: {len(chunks)}")
    print(f"  Synthesis time: {total_time:.1f}s ({total_time/60:.1f} min)")
    print(f"  Realtime factor: {total_time / total_duration:.1f}x")


if __name__ == '__main__':
    main()
