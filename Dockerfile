FROM python:3.12.10-slim

WORKDIR /app

# LibreOffice（PDF変換用）を追加インストール
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-core \
    fonts-noto-cjk \
    locales \
    curl \
    ca-certificates \
    && localedef -i ja_JP -c -f UTF-8 -A /usr/share/locale/locale.alias ja_JP.UTF-8 \
    && rm -rf /var/lib/apt/lists/*

ENV LANG=ja_JP.UTF-8
ENV LC_ALL=ja_JP.UTF-8

# 依存ライブラリのインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーション本体をコピー
COPY . .

# ================================
# Litestream の追加（ここが重要）
# ================================

# Litestream バイナリを追加
ADD https://github.com/benbjohnson/litestream/releases/download/v0.3.13/litestream-v0.3.13-linux-amd64.tar.gz /tmp/litestream.tar.gz
RUN tar -xzf /tmp/litestream.tar.gz -C /usr/local/bin && rm /tmp/litestream.tar.gz

# Litestream 設定ファイル
COPY litestream.yml /etc/litestream.yml

# エントリーポイント（DB restore + replicate）
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# SQLite 保存用ディレクトリ（Cloud Run のローカルディスク）
RUN mkdir -p /data

# ================================
# Streamlit 起動は entrypoint.sh に任せる
# ================================
ENTRYPOINT ["/entrypoint.sh"]

EXPOSE 8080
