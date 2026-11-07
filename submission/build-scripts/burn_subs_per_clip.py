#!/usr/bin/env python3
"""Burn subtitle PNGs into each clip, then concat and mux the Qwen narration."""
import json
import subprocess
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
VIDEO = ROOT / "build" / "video"
QWEN = VIDEO / "qwen"
OUT = ROOT / "04-演示视频.mp4"

VISUAL = {1: 8, 2: 10, 3: 17, 4: 12, 5: 11, 6: 14, 7: 20, 8: 13, 9: 5, 10: 9, 11: 6}
SEG_TO_CLIP = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 10, 10: 11}


def get_plan():
    timeline = json.loads((QWEN / "timeline.json").read_text(encoding="utf-8"))
    clip = dict(VISUAL)
    for seg in timeline["segments"]:
        ci = SEG_TO_CLIP[int(seg["id"])]
        clip[ci] = round(max(seg["audio_ms"] / 1000 + 1.5, VISUAL[ci]), 3)
    starts = {}
    t = 0.0
    for ci in range(1, 12):
        starts[ci] = round(t, 3)
        t += clip[ci]
    return clip, starts, round(t, 3)


def clip_of(sub):
    return 9 if sub["segment"] == "b-roll" else SEG_TO_CLIP[int(sub["segment"])]


def main():
    clip, starts, total = get_plan()
    subs = json.loads((QWEN / "subs.json").read_text(encoding="utf-8"))
    outs = []
    for ci in range(1, 12):
        src = VIDEO / f"qclip-{ci:02d}.mp4"
        dst = VIDEO / f"qsub-{ci:02d}.mp4"
        used = []
        for s in subs:
            if clip_of(s) != ci:
                continue
            rs = max(0.0, s["start"] - starts[ci])
            re = min(clip[ci], s["end"] - starts[ci])
            if re > rs:
                used.append((s, rs, re))
        if not used:
            subprocess.run(["cp", str(src), str(dst)], check=True)
            outs.append(dst)
            continue
        inputs = ["-i", str(src)]
        for s, _, _ in used:
            inputs += ["-loop", "1", "-t", str(clip[ci]), "-i", s["png"]]
        filters = []
        prev = "0:v"
        for i, (s, rs, re) in enumerate(used, start=1):
            out = f"v{i}"
            filters.append(
                f"[{prev}][{i}:v]overlay=0:0:enable='between(t,{rs:.3f},{re:.3f})'[{out}]"
            )
            prev = out
        subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error"] + inputs
            + ["-filter_complex", ";".join(filters), "-map", f"[{prev}]",
               "-c:v", "libx264", "-crf", "18", "-preset", "medium",
               "-pix_fmt", "yuv420p", "-r", "30", "-an", str(dst)],
            check=True,
        )
        outs.append(dst)
        print("burned", ci, "subs", len(used))

    with open(VIDEO / "qsubs.txt", "w") as fh:
        for out in outs:
            fh.write(f"file '{out.name}'\n")
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error", "-f", "concat", "-safe", "0",
         "-i", str(VIDEO / "qsubs.txt"), "-c", "copy", str(VIDEO / "qsub_video.mp4")],
        check=True,
    )
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error", "-i", str(VIDEO / "qsub_video.mp4"),
         "-i", str(QWEN / "narration.wav"), "-c:v", "copy", "-c:a", "aac",
         "-b:a", "192k", "-af", "volume=2.5dB", "-shortest", str(OUT)],
        check=True,
    )
    print("wrote", OUT, "total", total)


if __name__ == "__main__":
    main()
