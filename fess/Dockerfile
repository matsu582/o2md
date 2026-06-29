FROM ghcr.io/codelibs/fess:15.7.0

USER root

# Python, pip, LibreOffice, Tesseract のインストール
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    python3-venv \
    libreoffice \
    tesseract-ocr \
    tesseract-ocr-jpn \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

# o2md のインストール
RUN pip3 install --break-system-packages o2md

# extractor.xml を配置
COPY extractor.xml /opt/fess/app/WEB-INF/classes/crawler/extractor.xml

USER fess
