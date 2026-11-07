#!/usr/bin/env python3
"""Map H1-H3 headings to PDF page numbers, for the static TOC."""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, ".."))
MD = os.path.join(ROOT, "02-项目文档.md")
PDF = os.path.join(ROOT, "02-项目文档.pdf")
OUT = os.path.join(ROOT, "build", "page_map.json")


def normalize(value: str) -> str:
    value = value.replace("`", "").replace("**", "")
    return re.sub(r"\s+", "", value)


def main() -> None:
    text = subprocess.run(
        ["pdftotext", "-layout", PDF, "-"], check=True, capture_output=True, text=True
    ).stdout
    pages = text.split("\f")
    # cover + TOC occupy the first two pages
    body_start = 2  # zero-based index of page 3
    headings = []
    for line in open(MD, encoding="utf-8"):
        m = re.match(r"^(#{1,3})\s+(.*)$", line.strip())
        if m:
            headings.append(m.group(2).strip())
    page_map = {}
    for heading in headings:
        needle = normalize(heading)
        found = None
        for idx in range(body_start, len(pages)):
            if needle and needle in normalize(pages[idx]):
                found = idx + 1
                break
        if found is None:
            # tolerate a heading split by a line break: match a 10-char prefix
            short = needle[:10]
            for idx in range(body_start, len(pages)):
                if short and short in normalize(pages[idx]):
                    found = idx + 1
                    break
        page_map[heading] = found if found is not None else "??"
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(page_map, fh, ensure_ascii=False, indent=2)
    missing = [k for k, v in page_map.items() if v == "??"]
    print(f"mapped {len(page_map) - len(missing)}/{len(page_map)} headings")
    if missing:
        print("missing:", missing)


if __name__ == "__main__":
    main()
