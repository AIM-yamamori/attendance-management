#!/bin/bash
set -e

echo "=== ENTRYPOINT START ==="
echo "PORT=${PORT:-8080}"

DB_PATH="${DB_PATH:-/tmp/app.db}"
GCS_REPLICA_URL="gcs://aim-kintai-sqlite-backup/app-db-backup"

echo "=== DB RESTORE CHECK ==="
if [ -f "$DB_PATH" ]; then
  echo "DB already exists locally, skipping restore"
else
  echo "Attempting restore from GCS..."
  litestream restore -if-replica-exists -o "$DB_PATH" "$GCS_REPLICA_URL" || echo "No existing backup found, starting fresh"
  echo "Restored DB last modified: $(stat -c '%y' "$DB_PATH" 2>/dev/null || echo 'N/A')"
fi

# --- SIGTERM受信時の後処理を定義 ---
cleanup() {
  echo "=== SIGTERM RECEIVED: flushing litestream before shutdown ==="
  # litestream（配下のStreamlitごと）を止める
  if [ -n "$LITESTREAM_PID" ]; then
    kill -TERM "$LITESTREAM_PID" 2>/dev/null || true
    wait "$LITESTREAM_PID" 2>/dev/null || true
  fi
  echo "=== CLEANUP DONE ==="
  exit 0
}
trap cleanup SIGTERM

echo "=== STREAMLIT START (with Litestream replication) ==="
litestream replicate -exec "streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0" &
LITESTREAM_PID=$!
wait "$LITESTREAM_PID"