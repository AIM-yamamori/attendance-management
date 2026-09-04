#!/bin/bash
set -e

echo "=== ENTRYPOINT START ==="
echo "PORT=${PORT:-8080}"
echo "=== STREAMLIT START ==="

exec streamlit run app.py \
    --server.port="${PORT:-8080}" \
    --server.address=0.0.0.0