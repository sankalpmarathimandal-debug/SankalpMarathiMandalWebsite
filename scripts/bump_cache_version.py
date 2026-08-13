#!/usr/bin/env python3
"""Update the ?v= cache-busting query on style.css and main.js links in every
page, using the current commit's short SHA.

Runs automatically via GitHub Actions whenever style.css or main.js changes,
so visitors always get the latest CSS/JS instead of a stale cached copy.
"""
import glob
import re
import subprocess

def current_version():
    out = subprocess.run(["git", "rev-parse", "--short", "HEAD"],
                          capture_output=True, text=True).stdout.strip()
    return out or "1"

def main():
    ver = current_version()
    changed = []
    for path in glob.glob("*.html"):
        html = open(path, encoding="utf-8").read()
        new = re.sub(r'(href="assets/css/style\.css)(\?v=[^"]*)?(")', rf'\1?v={ver}\3', html)
        new = re.sub(r'(src="assets/js/main\.js)(\?v=[^"]*)?(")', rf'\1?v={ver}\3', new)
        if new != html:
            open(path, "w", encoding="utf-8").write(new)
            changed.append(path)
    print(f"Cache version set to {ver} in: {', '.join(changed) if changed else '(no changes needed)'}")

if __name__ == "__main__":
    main()
