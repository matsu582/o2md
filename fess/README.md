# Fess + o2md-filter 組み込み手順

Fessの`CommandExtractor`を利用して、`o2md-filter`をテキスト抽出エンジンとして組み込む手順です。
Javaコード不要、XML設定ファイルの追加のみで動作します。

## 構成概要

```
Fess クローラー
  → ファイル検出（MIMEタイプ判定）
    → CommandExtractor（o2md-filter $INPUT_FILE）
      → stdout からテキスト取得
        → OpenSearch にインデックス登録
```

## ネイティブ環境での組み込み

### 1. o2md-filter のインストール

```bash
pip install o2md

# LibreOffice（.doc/.xls/.ppt 変換に必要）
apt-get install libreoffice

# Tesseract OCR（画像/スキャンPDFのテキスト抽出に必要）
apt-get install tesseract-ocr tesseract-ocr-jpn tesseract-ocr-eng
```

### 2. extractor.xml の配置

Fessの`extractor.xml`を設定ディレクトリにコピー:

```bash
cp /opt/fess/lib/fess-crawler-*.jar内のcrawler/extractor.xml \
   /opt/fess/app/WEB-INF/classes/crawler/extractor.xml
```

または本リポジトリの `fess/extractor.xml` を配置:

```bash
cp fess/extractor.xml /opt/fess/app/WEB-INF/classes/crawler/extractor.xml
```

### 3. Fess を再起動

```bash
systemctl restart fess
```

## Docker 環境での組み込み

### カスタム Dockerfile を使用

`fess/Dockerfile` を使用してo2md-filterを含むカスタムFessイメージをビルドします。

```bash
cd fess
docker build -t fess-o2md .
```

`compose.yaml` で使用:

```yaml
services:
  fess:
    image: fess-o2md
    # ... 他の設定
```

### Docker Compose 完全構成例

```bash
# OpenSearch起動に必要なカーネル設定
sudo sysctl -w vm.max_map_count=262144

cd fess
docker compose up -d
```

## 対象MIMEタイプ

| MIMEタイプ | ファイル形式 |
| --- | --- |
| `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` | .xlsx |
| `application/vnd.ms-excel` | .xls |
| `application/vnd.openxmlformats-officedocument.wordprocessingml.document` | .docx |
| `application/msword` | .doc |
| `application/vnd.openxmlformats-officedocument.presentationml.presentation` | .pptx |
| `application/vnd.ms-powerpoint` | .ppt |
| `application/pdf` | .pdf |
| `image/jpeg` | .jpg |
| `image/png` | .png |
| `image/webp` | .webp |
| `image/gif` | .gif |
| `image/bmp` | .bmp |
| `image/tiff` | .tiff |

## 注意事項

- `o2md-filter` はファイルパスを受け取りstdoutにプレーンテキストを出力します
- LibreOfficeがない場合、`.doc`/`.xls`/`.ppt` の変換はエラーになります
- Tesseractがない場合、画像/スキャンPDFのOCRは利用できません
- OCRエンジンをmanga-ocrに変更する場合: `o2md-filter --ocr-engine manga-ocr $INPUT_FILE`
- `compose.yaml`はローカル開発・検証用です。本番環境ではOpenSearchのセキュリティプラグインを有効化してください
