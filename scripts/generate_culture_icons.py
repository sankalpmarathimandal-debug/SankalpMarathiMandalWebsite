#!/usr/bin/env python3
"""Scan assets/images/culture-icons/ and write data/culture-icons.json.

Runs automatically via GitHub Actions whenever images are added, renamed, or
removed in that folder. No coding needed to change the homepage culture
icon ribbon — just drop an image file into the folder (or use the admin
panel) and this regenerates the manifest the site reads. No captions are
shown for these icons, so file names don't matter for display — only for
your own organization. Newest images (by git history) are listed first.
"""
import json, os, subprocess

FOLDER = "assets/images/culture-icons"
OUT = "data/culture-icons.json"
EXTS = (".jpg", ".jpeg", ".png", ".webp", ".gif", ".svg")

def commit_time(path):
    try:
        t = subprocess.run(["git", "log", "-1", "--format=%ct", "--", path],
                           capture_output=True, text=True).stdout.strip()
        return int(t) if t else 0
    except Exception:
        return 0

files = [f for f in os.listdir(FOLDER)
         if f.lower().endswith(EXTS) and not f.startswith(".")]
files.sort(key=lambda f: (-commit_time(os.path.join(FOLDER, f)), f))

items = [{"src": f"{FOLDER}/{f}"} for f in files]
os.makedirs("data", exist_ok=True)
with open(OUT, "w", encoding="utf-8") as fh:
    json.dump(items, fh, indent=2, ensure_ascii=False)
print(f"Wrote {OUT} with {len(items)} culture icons")
