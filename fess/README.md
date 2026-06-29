# Fess + o2md-filter 組み込み手順

Fessの`CommandExtractor`を利用して、`o2md-filter`をテキスト抽出エンジンとして組み込む手順です。
Javaコード不要、XML設定ファイルの追加のみで動作します。

## 構成概要

```
Fess クローラー
  → ファイル検出（MIMEタイプ判定）
    → CommandExtractor（o2md-filter $INPUT_FILE $OUTPUT_FILE）
      → OpenSearch にインデックス登録
```

`o2md-filter`は第2引数に出力ファイルパスを指定できるため、
CommandExtractorから直接呼び出せます。

## ネイティブ環境での組み込み

### 1. o2md-filter のインストール

```bash
pip install o2md

# LibreOffice（.doc/.xls/.ppt 変換に必要）
apt-get install libreoffice

# Tesseract OCR（画像/スキャンPDFのテキスト抽出に必要）
apt-get install tesseract-ocr tesseract-ocr-jpn tesseract-ocr-eng
```

### 2. XML設定ファイルの配置

```bash
# extractor.xml（MIMEタイプ→o2md-filter マッピング）
cp fess/extractor.xml /opt/fess/app/WEB-INF/classes/crawler/extractor.xml

# contentlength.xml（大容量ファイル対応、デフォルト10MB→50MB）
cp fess/contentlength.xml /opt/fess/app/WEB-INF/classes/crawler/contentlength.xml
```

### 3. Fess を再起動

```bash
systemctl restart fess
```

## Docker 環境での組み込み

### 起動手順

```bash
# OpenSearch起動に必要なカーネル設定
sudo sysctl -w vm.max_map_count=262144

cd fess
# compose.yaml内の /path/to/documents を実際のパスに変更
docker compose up -d
```

### 設定ファイル

| ファイル | 役割 |
| --- | --- |
| `Dockerfile` | o2md-filter + LibreOffice + Tesseract入りカスタムFessイメージ |
| `extractor.xml` | MIMEタイプごとのExtractor登録（postConstruct形式） |
| `contentlength.xml` | ContentLengthHelperの上限を50MBに拡大 |
| `compose.yaml` | Fess + OpenSearch のDocker Compose構成 |

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
| `application/x-js-taro` | .jtd（一太郎） |
| `application/octet-stream` | 不明な形式（o2md-filterが自動判別） |
| `image/jpeg`, `image/png`, `image/webp`, `image/tiff`, `image/bmp`, `image/gif` | 画像（OCR） |

## 注意事項

- `o2md-filter`は第2引数で出力ファイルを指定できます（`o2md-filter input.pdf output.txt`）。省略時はstdoutに出力します
- Fess 15.7の`extractor.xml`は`postConstruct`形式でExtractorを登録する必要があります（`<property name="extractorMap">`形式は不可）
- LibreOfficeがない場合、`.doc`/`.xls`/`.ppt` の変換はエラーになります
- Tesseractがない場合、画像/スキャンPDFのOCRは利用できません
- デフォルトのファイルサイズ上限は10MBです。大容量ファイルを扱う場合は`contentlength.xml`で`defaultMaxLength`を変更してください
- `compose.yaml`はローカル開発・検証用です。本番環境ではOpenSearchのセキュリティプラグインを有効化してください
