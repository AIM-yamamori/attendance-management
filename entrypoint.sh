# ジョブモード
if [ "$JOB_MODE" = "shell" ]; then
  exec /bin/bash
fi

#!/bin/bash
set -e

echo "=== ENTRYPOINT START ==="
echo "PORT=${PORT:-8080}"

DB_PATH="${DB_PATH:-/tmp/app.db}"

echo "=== DB RESTORE CHECK ==="
if [ -f "$DB_PATH" ]; then
  echo "DB already exists locally, skipping restore"
else
  echo "Attempting restore from GCS..."
  litestream restore -if-replica-exists -o "$DB_PATH" "gcs://aim-kintai-sqlite-backup/app-db-backup" || echo "No existing backup found, starting fresh"
fi

echo "=== STREAMLIT START (with Litestream replication) ==="
exec litestream replicate -exec "streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0"