#!/usr/bin/env bash
set -euo pipefail

CHROME_BIN="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
DEBUG_PROFILE="${TOPCV_CHROME_DEBUG_PROFILE:-/tmp/topcv-chrome-debug}"
DEBUG_PORT="${TOPCV_CHROME_DEBUG_PORT:-9222}"

exec "$CHROME_BIN" \
  --remote-debugging-port="$DEBUG_PORT" \
  --user-data-dir="$DEBUG_PROFILE"
