#!/usr/bin/env python3
"""Build the demo video with Qwen TTS narration and burned-in subtitles.

Inputs:
  build/video/qwen/timeline.json   (from tts_qwen.py)
  build/video/slides/*.png         (PPT pages)
  build/video/ui-demo.mp4          (20 s UI demo)
Output:
  04-演示视频.mp4
"""
import json
import math
import subprocess
import sys
import wave
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
VIDEO = ROOT / "build" / "video"
QWEN = VIDEO / "qwen"
SLIDES = VIDEO / "slides"
SUBS = VIDEO / "subs"
OUT = ROOT / "04-演示视频.mp4"

VISUAL = {1: 8, 2: 10, 3: 17, 4: 12, 5: 11, 6: 14, 7: 20, 8: 13, 9: 5, 10: 9, 11: 6}
SEG_TO_CLIP = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 10, 10: 11}
SLIDE_MAP = {1: "slide-01", 2: "slide-03", 3: "slide-04", 4: "slide-06", 5: "slide-04", 6: "slide-07", 8: "slide-11", 9: "slide-12", 10: "slide-13", 11: "slide-16"}
FONT_PATH = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"


def get_durations():
    timeline = json.loads((QWEN / "timeline.json").read_text(encoding="utf-8"))
    clip = dict(VISUAL)
    for seg in timeline["segments"]:
        idx = int(seg["id"])
        ci = SEG_TO_CLIP[idx]
        clip[ci] = round(max(seg["audio_ms"] / 1000 + 1.5, VISUAL[ci]), 3)
    starts = {}
    t = 0.0
    for ci in range(1, 12):
        starts[ci] = round(t, 3)
        t += clip[ci]
    return timeline, clip, starts, round(t, 3)


def groups_for_segment(seg):
    cues = seg["cues"]
    if len(cues) == 1:
        return [(cues, 0.0, cues[-1]["end_ms"] / 1000, cues[0]["text"])]
    mid = len(cues) // 2
    g1, g2 = cues[:mid], cues[mid:]
    text1 = "".join(c["text"] for c in g1)
    text2 = "".join(c["text"] for c in g2)
    return [
        (g1, 0.0, g1[-1]["end_ms"] / 1000, text1),
        (g2, g2[0]["start_ms"] / 1000, g2[-1]["end_ms"] / 1000, text2),
    ]


def make_subtitle(text, path, font_size=34):
    font = ImageFont.truetype(FONT_PATH, font_size)
    img = Image.new("RGBA", (1920, 1080), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # wrap by character width
    max_w = 1500
    lines = []
    cur = ""
    for ch in text:
        trial = cur + ch
        if draw.textbbox((0, 0), trial, font=font)[2] <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    if len(lines) > 3:
        font = ImageFont.truetype(FONT_PATH, 30)
        draw = ImageDraw.Draw(img)
        lines = []
        cur = ""
        for ch in text:
            trial = cur + ch
            if draw.textbbox((0, 0), trial, font=font)[2] <= max_w:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    line_h = font_size + 14
    pad_x, pad_y = 28, 18
    text_w = max(draw.textbbox((0, 0), line, font=font)[2] for line in lines)
    bar_w = min(1800, text_w + pad_x * 2)
    bar_h = len(lines) * line_h + pad_y * 2
    x0 = (1920 - bar_w) // 2
    y0 = 1080 - 42 - bar_h
    draw.rounded_rectangle([x0, y0, x0 + bar_w, y0 + bar_h], radius=18,
                           fill=(0, 0, 0, 185), outline=(255, 255, 255, 35), width=1)
    y = y0 + pad_y
    for line in lines:
        w = draw.textbbox((0, 0), line, font=font)[2]
        x = (1920 - w) // 2
        draw.text((x + 2, y + 2), line, font=font, fill=(0, 0, 0, 160))
        draw.text((x, y), line, font=font, fill=(255, 255, 255, 255))
        y += line_h
    img.save(path)


def build_subtitles(timeline, clip_starts, clip_durations):
    SUBS.mkdir(parents=True, exist_ok=True)
    subs = []
    for seg in timeline["segments"]:
        idx = int(seg["id"])
        ci = SEG_TO_CLIP[idx]
        base = clip_starts[ci]
        for group, gs, ge, text in groups_for_segment(seg):
            start = base + gs
            end = min(base + ge + 0.10, base + clip_durations[ci] - 0.08)
            subs.append({"text": text, "start": round(start, 3), "end": round(end, 3), "segment": seg["id"]})
    # B-roll subtitle for the ASE page
    b0 = clip_starts[9]
    subs.append({"text": "界面遮挡研究 · 真实应用验证", "start": round(b0, 3),
                 "end": round(b0 + VISUAL[9], 3), "segment": "b-roll"})
    subs.sort(key=lambda s: s["start"])
    for i, s in enumerate(subs, start=1):
        path = SUBS / f"cue-{i:02d}.png"
        make_subtitle(s["text"], path)
        s["png"] = str(path)
        print("sub", i, round(s["start"], 2), round(s["end"], 2), s["text"][:40])
    (QWEN / "subs.json").write_text(json.dumps(subs, ensure_ascii=False, indent=2), encoding="utf-8")
    return subs


def run(cmd, **kwargs):
    subprocess.run(cmd, check=True, **kwargs)


def build_clips(clip_durations):
    clips = []
    for ci in range(1, 12):
        dur = clip_durations[ci]
        out = VIDEO / f"qclip-{ci:02d}.mp4"
        if ci == 7:
            run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(VIDEO / "ui-demo.mp4"),
                 "-t", str(dur), "-c:v", "libx264", "-crf", "18", "-preset", "medium",
                 "-pix_fmt", "yuv420p", "-r", "30", "-an", str(out)])
        else:
            name = SLIDE_MAP[ci]
            src = SLIDES / f"{name}.png"
            frames = int(round(dur * 30))
            vf = (f"scale=1920:1080,zoompan=z='min(zoom+0.00035,1.045)':"
                  f"x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':d={frames}:"
                  f"s=1920x1080:fps=30,format=yuv420p")
            run(["ffmpeg", "-y", "-loglevel", "error", "-loop", "1", "-i", str(src), "-vf", vf,
                 "-t", str(dur), "-r", "30", "-c:v", "libx264", "-crf", "18", "-preset", "medium",
                 "-an", str(out)])
        clips.append(out)
    with open(VIDEO / "qclips.txt", "w") as fh:
        for c in clips:
            fh.write(f"file '{c.name}'\n")
    run(["ffmpeg", "-y", "-loglevel", "error", "-f", "concat", "-safe", "0",
         "-i", str(VIDEO / "qclips.txt"), "-c", "copy", str(VIDEO / "qvideo_noaudio.mp4")])


def build_narration(timeline, clip_starts, total):
    inputs = ["-f", "lavfi", "-i", f"anullsrc=r=48000:cl=mono:d={total}"]
    for seg in timeline["segments"]:
        inputs += ["-i", str(QWEN / "segments" / f"{seg['id']}.wav")]
    filters = []
    for i, seg in enumerate(timeline["segments"], start=1):
        ci = SEG_TO_CLIP[int(seg["id"])]
        off = int(round(clip_starts[ci] * 1000))
        filters.append(f"[{i}:a]adelay={off}:all=1[a{i}]")
    mix = "[0:a]" + "".join(f"[a{i}]" for i in range(1, len(timeline["segments"]) + 1))
    filters.append(f"{mix}amix=inputs={len(timeline['segments']) + 1}:normalize=0[out]")
    run(["ffmpeg", "-y", "-loglevel", "error"] + inputs + [
        "-filter_complex", ";".join(filters), "-map", "[out]", "-c:a", "pcm_s16le",
        str(QWEN / "narration.wav")])


def render(subs, total):
    inputs = ["-i", str(VIDEO / "qvideo_noaudio.mp4")]
    for s in subs:
        inputs += ["-loop", "1", "-t", str(total), "-i", s["png"]]
    audio_idx = len(subs) + 1
    inputs += ["-i", str(QWEN / "narration.wav")]
    filters = []
    prev = "0:v"
    for i, s in enumerate(subs, start=1):
        out = f"v{i}"
        filters.append(f"[{prev}][{i}:v]overlay=0:0:enable='between(t,{s['start']:.3f},{s['end']:.3f})'[{out}]")
        prev = out
    cmd = ["ffmpeg", "-y", "-loglevel", "error"] + inputs + [
        "-filter_complex", ";".join(filters), "-map", f"[{prev}]", "-map", f"{audio_idx}:a",
        "-c:v", "libx264", "-crf", "18", "-preset", "medium", "-pix_fmt", "yuv420p",
        "-r", "30", "-c:a", "aac", "-b:a", "192k", "-shortest", str(OUT)]
    run(cmd)


def main():
    timeline, clip_durations, clip_starts, total = get_durations()
    print("clip_durations", clip_durations)
    print("clip_starts", clip_starts)
    print("total", total)
    subs = build_subtitles(timeline, clip_starts, clip_durations)
    print("subtitles", len(subs))
    build_clips(clip_durations)
    build_narration(timeline, clip_starts, total)
    subprocess.run([sys.executable, str(HERE / "burn_subs_per_clip.py")], check=True)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
