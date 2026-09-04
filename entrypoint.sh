#!/bin/bash
set -e

# 起動時: GCSに既存バックアップがあれば復元
if [ -f /data/app.db ]; then
  echo "Local DB exists, skipping restore"
else
  litestream restore -if-replica-exists -o /data/app.db "gcs://aim-kintai-sqlite-backup/app-db-backup"
fi

# Litestreamでレプリケーションしながらアプリを起動
exec litestream replicate -exec "streamlit run app.py --server.port=8080 --server.address=0.0.0.0"