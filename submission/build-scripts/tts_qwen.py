#!/usr/bin/env python3
"""Generate the demo narration with Qwen TTS and per-sentence subtitles.

The API key is read from DASHSCOPE_API_KEY or /tmp/dashscope_key; it is never
written into the repository. Outputs go to submission/build/video/qwen/.
"""
import json
import os
import re
import subprocess
import time
import wave
from pathlib import Path

import dashscope
from dashscope.audio.tts_v2 import SpeechSynthesizer

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SCRIPT = ROOT / "build-scripts" / "video_script.json"
OUT = ROOT / "build" / "video" / "qwen"
CUES = OUT / "cues"
SEGS = OUT / "segments"

MODEL = "qwen-audio-3.1-tts-flash"
VOICE = "longshu_v3.1"
BASE_URL = "wss://cn-beijing.maas.qianwenaiapi.com/api-ws/v1/inference"
GAP_MS = 180


def load_key() -> str:
    key = os.environ.get("DASHSCOPE_API_KEY", "").strip()
    if not key:
        path = Path("/tmp/dashscope_key")
        if path.exists():
            key = path.read_text(encoding="utf-8").strip()
    if not key:
        raise SystemExit("DASHSCOPE_API_KEY is missing")
    return key


def split_cues(text: str) -> list[str]:
    text = re.sub(r"\[\[slnc \d+\]\]", "", text).strip()
    parts = re.split(r"(?<=[。；！？])", text)
    out = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if len(part) <= 36:
            out.append(part)
            continue
        subs = re.split(r"(?<=[，、；])", part)
        cur = ""
        for sub in subs:
            if not sub:
                continue
            if len(cur) + len(sub) <= 36:
                cur += sub
            else:
                if cur:
                    out.append(cur)
                cur = sub
        if cur:
            out.append(cur)
    final = []
    for cue in out:
        while len(cue) > 50:
            final.append(cue[:46])
            cue = cue[46:]
        if cue:
            final.append(cue)
    return final


def synth(text: str, key: str) -> bytes:
    dashscope.api_key = key
    dashscope.base_websocket_api_url = BASE_URL
    last = None
    for attempt in range(4):
        try:
            synthesizer = SpeechSynthesizer(
                model=MODEL, voice=VOICE, speech_rate=0.98, pitch_rate=1.0,
                instruction="请用自然、沉稳的新闻播报语气，像面对面讲解，语速适中，减少机械感。")
            audio = synthesizer.call(text)
            if audio:
                return audio
            last = "empty audio"
        except Exception as exc:  # API/network errors are retried
            last = f"{type(exc).__name__}: {exc}"
        time.sleep(0.8 * (attempt + 1))
    raise RuntimeError(f"TTS failed for {text!r}: {last}")


def mp3_to_wav(src: Path, dst: Path) -> None:
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error", "-i", str(src), "-ar", "48000", "-ac", "1", str(dst)],
        check=True,
    )


def write_silence(wav: wave.Wave_write, ms: int) -> None:
    frames = int(48000 * ms / 1000)
    wav.writeframes(b"\x00\x00" * frames)


def main() -> None:
    key = load_key()
    CUES.mkdir(parents=True, exist_ok=True)
    SEGS.mkdir(parents=True, exist_ok=True)
    script = json.loads(SCRIPT.read_text(encoding="utf-8"))
    all_cues = []
    segments = []
    for seg in script:
        seg_id = seg["id"]
        cues = split_cues(seg["narration"])
        seg_wav = SEGS / f"{seg_id}.wav"
        cue_rows = []
        with wave.open(str(seg_wav), "wb") as wav:
            wav.setnchannels(1)
            wav.setsampwidth(2)
            wav.setframerate(48000)
            cursor_ms = 0
            for idx, cue in enumerate(cues, start=1):
                mp3 = CUES / f"{seg_id}-{idx:02d}.mp3"
                wav_cue = CUES / f"{seg_id}-{idx:02d}.wav"
                if not mp3.exists() or mp3.stat().st_size == 0:
                    mp3.write_bytes(synth(cue, key))
                mp3_to_wav(mp3, wav_cue)
                with wave.open(str(wav_cue), "rb") as src:
                    data = src.readframes(src.getnframes())
                    duration_ms = round(len(data) / 2 / 48000 * 1000)
                wav.writeframes(data)
                start_ms = cursor_ms
                end_ms = start_ms + duration_ms
                cue_rows.append({
                    "segment": seg_id,
                    "index": idx,
                    "text": cue,
                    "start_ms": start_ms,
                    "end_ms": end_ms,
                })
                all_cues.append(cue_rows[-1])
                cursor_ms = end_ms
                if idx < len(cues):
                    write_silence(wav, GAP_MS)
                    cursor_ms += GAP_MS
        segments.append({
            "id": seg_id,
            "audio_ms": cursor_ms,
            "cues": cue_rows,
            "image": seg.get("image"),
        })
        print(seg_id, "cues", len(cues), "audio", round(cursor_ms / 1000, 2), "s")

    timeline = {"model": MODEL, "voice": VOICE, "segments": segments, "cues": all_cues}
    (OUT / "timeline.json").write_text(json.dumps(timeline, ensure_ascii=False, indent=2), encoding="utf-8")
    print("wrote", OUT / "timeline.json")


if __name__ == "__main__":
    main()
