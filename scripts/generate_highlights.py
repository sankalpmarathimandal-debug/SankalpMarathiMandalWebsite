#!/usr/bin/env python3
"""Scan assets/images/highlights/ and write data/highlights.json.

Runs automatically via GitHub Actions whenever images are added or removed.
Title rule: 'gudi-padwa-01.jpg' -> 'Gudi Padwa'; files named 'highlight-*',
'img*' or random IDs get no caption.
"""
import json, os, re, subprocess

FOLDER = "assets/images/highlights"
OUT = "data/highlights.json"
EXTS = (".jpg", ".jpeg", ".png", ".webp", ".gif")

def commit_time(path):
    try:
        t = subprocess.run(["git", "log", "-1", "--format=%ct", "--", path],
                           capture_output=True, text=True).stdout.strip()
        return int(t) if t else 0
    except Exception:
        return 0

def title_for(name):
    base = os.path.splitext(name)[0]
    base = re.sub(r"[-_ ]?\d+$", "", base)          # strip trailing numbers
    if re.fullmatch(r"(highlight|img|image|photo|pic)[-_ ]?", base, re.I) or not base:
        return ""
    if re.fullmatch(r"[0-9a-f-]{16,}(_\d+_\w+)?", base, re.I):  # camera/UUID names
        return ""
    return base.replace("-", " ").replace("_", " ").strip().title()

files = [f for f in os.listdir(FOLDER)
         if f.lower().endswith(EXTS) and not f.startswith(".")]
files.sort(key=lambda f: (-commit_time(os.path.join(FOLDER, f)), f))

items = [{"photo": f"{FOLDER}/{f}", "title": title_for(f)} for f in files]
os.makedirs("data", exist_ok=True)
with open(OUT, "w", encoding="utf-8") as fh:
    json.dump(items, fh, indent=2, ensure_ascii=False)
print(f"Wrote {OUT} with {len(items)} photos")
