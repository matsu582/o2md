FROM ghcr.io/codelibs/fess:15.7.0

USER root

# Python + OCR/LibreOffice
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
# 旧一太郎(ver4-7)のテキスト抽出は 0.1.3 以降が必要
ARG O2MD_VERSION=0.1.3
RUN python3 -m venv --system-site-packages /opt/o2md-venv \
    && /opt/o2md-venv/bin/pip install --no-cache-dir --no-deps o2md==${O2MD_VERSION} \
    && /opt/o2md-venv/bin/pip install --no-cache-dir \
       olefile openpyxl pdfplumber pymupdf pytesseract \
       python-docx python-pptx requests pdfminer.six \
       "Pillow>=12.0.0" "numpy>=2.0.0"

ENV PATH="/opt/o2md-venv/bin:${PATH}"

# fess-webapp-mcp プラグイン（MCPサーバー機能）
# https://github.com/codelibs/fess-webapp-mcp
ARG FESS_MCP_VERSION=15.7.0
ADD https://repo1.maven.org/maven2/org/codelibs/fess/fess-webapp-mcp/${FESS_MCP_VERSION}/fess-webapp-mcp-${FESS_MCP_VERSION}.jar \
    /usr/share/fess/app/WEB-INF/lib/fess-webapp-mcp-${FESS_MCP_VERSION}.jar

# Fessクローラー設定の上書き
COPY extractor.xml /usr/share/fess/app/WEB-INF/classes/crawler/extractor.xml
COPY contentlength.xml /usr/share/fess/app/WEB-INF/classes/crawler/contentlength.xml

# 旧一太郎拡張子のMIMEタイプ定義（Tika用）
COPY custom-mimetypes.xml /usr/share/fess/app/WEB-INF/classes/org/apache/tika/mime/custom-mimetypes.xml
