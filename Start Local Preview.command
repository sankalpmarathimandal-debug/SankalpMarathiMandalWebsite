#!/bin/bash
# Sankalp Website — local preview server
# Double-click this file to preview the site with all data loading correctly.
cd "$(dirname "$0")"
python3 scripts/generate_highlights.py 2>/dev/null
PORT=8765
echo "Starting Sankalp website preview at http://localhost:$PORT"
echo "Press Ctrl+C in this window to stop the server."
( sleep 1 && open "http://localhost:$PORT" ) &
python3 -m http.server $PORT
