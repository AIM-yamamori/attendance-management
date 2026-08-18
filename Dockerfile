FROM python:3.12.10-slim

WORKDIR /app

# LibreOffice（PDF変換用）を追加インストール
# --no-install-recommends で最小限の依存のみ導入し、イメージサイズを抑える
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-core \
    fonts-noto-cjk \
    locales \
    && localedef -i ja_JP -c -f UTF-8 -A /usr/share/locale/locale.alias ja_JP.UTF-8 \
    && rm -rf /var/lib/apt/lists/*

ENV LANG=ja_JP.UTF-8
ENV LC_ALL=ja_JP.UTF-8

# 依存ライブラリのインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーション本体をコピー
COPY . .

EXPOSE 8501

VOLUME ["/app/data"]

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]