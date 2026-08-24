#!/bin/bash
# ── Redacter Launcher ──────────────────────────────────────────────────────
# Starts the local Flask server and opens the browser.
# Usage: ./start.sh
# Stop:  Ctrl+C

set -e
cd "$(dirname "$0")"

# Check deps
python3 -c "import flask, fitz" 2>/dev/null || {
  echo ""
  echo "  ⚠  Missing dependencies. Run:"
  echo "     pip3 install flask pymupdf python-docx"
  echo ""
  exit 1
}

echo ""
echo "  ████████████████████████████████████████████████████"
echo "  █                                                  █"
echo "  █   REDACTER  —  Local PDF/Word Redaction Tool     █"
echo "  █                                                  █"
echo "  █   http://localhost:5050                          █"
echo "  █   Press Ctrl+C to stop                          █"
echo "  █                                                  █"
echo "  ████████████████████████████████████████████████████"
echo ""

# Start server
python3 app.py &
SERVER_PID=$!

# Give it a moment to start
sleep 1.5

# Open browser (macOS)
if command -v open &>/dev/null; then
  open http://localhost:5050
elif command -v xdg-open &>/dev/null; then
  xdg-open http://localhost:5050
fi

# Wait for server (Ctrl+C kills both)
trap "kill $SERVER_PID 2>/dev/null; echo '  Server stopped.'; exit 0" INT TERM
wait $SERVER_PID
