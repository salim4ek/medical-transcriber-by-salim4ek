#!/bin/bash
# Hourly cleanup: prune input/, output/, temp/ and truncate logs.
# Files newer than 30 min are skipped to protect in-flight processing.
set -u
cd "$(dirname "$0")"

STAMP=$(date +%FT%T)

IN=$(find input/        -type f -mmin +30 2>/dev/null | wc -l)
OUT=$(find output/       -type f -mmin +30 2>/dev/null | wc -l)
TMP=$(find temp/         -type f -mmin +30 2>/dev/null | wc -l)
find input/        -type f -mmin +30 -delete 2>/dev/null || true
find output/       -type f -mmin +30 -delete 2>/dev/null || true
find temp/         -type f -mmin +30 -delete 2>/dev/null || true

# Truncate logs > 500 KB to last 500 KB
for log in transcriber.log bot.log; do
  if [ -f "$log" ]; then
    size=$(stat -c%s "$log")
    if [ "$size" -gt 524288 ]; then
      tail -c 524288 "$log" > "$log.tmp" && mv "$log.tmp" "$log"
    fi
  fi
done

echo "[cleanup $STAMP] removed input:$IN output:$OUT temp:$TMP, logs truncated to ~500K"
