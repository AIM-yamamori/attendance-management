#!/bin/bash
set -e

echo "=== ENTRYPOINT START ==="
echo "DB_PATH=$DB_PATH"
echo "PORT=${PORT:-8080}"

DB_PATH="/mnt/disks/ephemeral/app.db"

echo "=== RESTORE START ==="

if [ -f "$DB_PATH" ]; then
  echo "Local DB exists, skipping restore"
else
  litestream restore \
    -if-replica-exists \
    -o "$DB_PATH" \
    "gcs://aim-kintai-sqlite-backup/app-db-backup"
fi

echo "=== RESTORE END ==="
echo "=== STREAMLIT START ==="

exec litestream replicate \
  -exec "streamlit run app.py --server.port=8080 --server.address=0.0.0.0"