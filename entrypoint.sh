#!/bin/bash
set -e

DB_PATH="/mnt/disks/ephemeral/app.db"

# 起動時: GCSに既存バックアップがあれば復元
if [ -f "$DB_PATH" ]; then
  echo "Local DB exists, skipping restore"
else
  litestream restore -if-replica-exists -o "$DB_PATH" "gcs://aim-kintai-sqlite-backup/app-db-backup"
fi

# Litestreamでレプリケーションしながらアプリを起動
exec litestream replicate -exec "streamlit run app.py --server.port=8080 --server.address=0.0.0.0"
