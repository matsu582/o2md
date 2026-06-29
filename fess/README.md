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

## Fess標準（Apache Tika）との比較

Fessはデフォルトで[Apache Tika](https://tika.apache.org/)をテキスト抽出エンジンとして使用します。
o2md-filterに置き換えることで対応形式が広がりますが、トレードオフもあります。

### 対応形式の比較

| ファイル形式 | Fess標準（Tika） | o2md-filter |
| --- | --- | --- |
| .xlsx / .docx / .pptx | テキスト抽出可能 | テキスト抽出可能 |
| .xls / .doc / .ppt | テキスト抽出可能 | テキスト抽出可能（LibreOffice経由） |
| .pdf（テキストベース） | テキスト抽出可能 | テキスト抽出可能 |
| .pdf（スキャン/画像ベース） | テキスト抽出不可 | OCRでテキスト抽出可能 |
| .jtd（一太郎） | テキスト抽出不可 | 独自パーサーでテキスト抽出可能 |
| 画像（jpg/png/webp等） | テキスト抽出不可 | OCRでテキスト抽出可能 |

### 処理速度

| 対象 | Fess標準（Tika） | o2md-filter |
| --- | --- | --- |
| Office文書（xlsx/docx/pptx） | 高速（Javaネイティブ、Apache POI） | やや遅い（Pythonプロセス起動のオーバーヘッド） |
| 旧形式（xls/doc/ppt） | 高速（Javaネイティブ） | 遅い（LibreOffice起動が必要） |
| PDF（テキストベース） | 高速（PDFBox） | やや遅い（pdfplumber/PyMuPDF） |
| PDF（スキャン/画像） | — | 遅い（OCR処理のため1ページ数秒） |
| 画像 | — | 遅い（OCR処理） |

### Fess標準（Tika）の特徴

- **高速**: Javaネイティブで動作し、外部プロセス起動のオーバーヘッドがない
- **メタデータ抽出が豊富**: 作成者、作成日、キーワード等のメタデータを自動抽出
- **対応形式が広い**: Office文書以外にもRTF、HTML、XML、電子メール等200以上の形式に対応
- **追加依存なし**: Fessに同梱されており追加インストール不要
- **安定性**: 長年の実績がある成熟したライブラリ

### o2md-filterの特徴

- **一太郎対応**: OLE2バイナリの独自パーサーによりTikaでは不可能な.jtdのテキスト抽出が可能
- **OCR対応**: スキャンPDF・画像からテキスト抽出可能（Tesseract/manga-ocr/sarashina選択可）
- **日本語最適化**: 日本語文書のテキスト抽出に特化した処理
- **処理速度が遅い**: ファイルごとにPythonプロセスを起動するため、大量ファイルのクロール時にTikaより時間がかかる
- **追加依存が多い**: Python、LibreOffice、Tesseract等の追加インストールが必要
- **Dockerイメージが大きい**: 上記依存によりカスタムイメージのサイズが増加（約1.5GB増）
- **メタデータ抽出なし**: テキスト本文のみを抽出し、ドキュメントメタデータは返さない

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

## MCPサーバー連携

[fess-webapp-mcp](https://github.com/codelibs/fess-webapp-mcp)プラグインにより、FessがMCPサーバーとして動作します。
Dockerfileにプラグインが含まれているため、追加設定なしで`POST /mcp`エンドポイントが利用可能です。

### MCPツール

| ツール | 説明 |
| --- | --- |
| `search` | キーワード検索（`q`, `num`, `start`, `sort`） |
| `suggest` | オートコンプリート |
| `get_document` | ドキュメントIDで個別取得 |
| `get_index_stats` | インデックス統計情報 |

### Claude Desktopとの連携

```json
{
  "mcpServers": {
    "fess": {
      "command": "npx",
      "args": ["-y", "mcp-remote", "http://localhost:8080/mcp"]
    }
  }
}
```

### 動作確認

```bash
# ping
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"ping","params":{}}'

# 検索
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"search","arguments":{"q":"検索キーワード","num":5}}}'
```

## 注意事項

- `o2md-filter`は第2引数で出力ファイルを指定できます（`o2md-filter input.pdf output.txt`）。省略時はstdoutに出力します
- Fess 15.7の`extractor.xml`は`postConstruct`形式でExtractorを登録する必要があります（`<property name="extractorMap">`形式は不可）
- LibreOfficeがない場合、`.doc`/`.xls`/`.ppt` の変換はエラーになります
- Tesseractがない場合、画像/スキャンPDFのOCRは利用できません
- デフォルトのファイルサイズ上限は10MBです。大容量ファイルを扱う場合は`contentlength.xml`で`defaultMaxLength`を変更してください
- `compose.yaml`はローカル開発・検証用です。本番環境ではOpenSearchのセキュリティプラグインを有効化してください
