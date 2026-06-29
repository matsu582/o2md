FROM ghcr.io/codelibs/fess:15.7.0

USER root

# Python ランタイム + OCR/LibreOffice + opencv
RUN apk add --no-cache \
    python3 \
    py3-pip \
    py3-virtualenv \
    py3-opencv \
    libreoffice \
    tesseract-ocr \
    tesseract-ocr-data-jpn \
    tesseract-ocr-data-eng

# venv作成（system-site-packagesでcv2を参照）+ o2md インストール
RUN python3 -m venv --system-site-packages /opt/o2md-venv \
    && /opt/o2md-venv/bin/pip install --no-cache-dir --no-deps o2md \
    && /opt/o2md-venv/bin/pip install --no-cache-dir \
       olefile openpyxl pdfplumber pymupdf pytesseract \
       python-docx python-pptx requests

# o2md-filter へのPATH設定
ENV PATH="/opt/o2md-venv/bin:${PATH}"

# extractor.xml を配置
COPY extractor.xml /opt/fess/app/WEB-INF/classes/crawler/extractor.xml
