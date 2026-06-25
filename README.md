# Office to Markdown Converter (o2md)

Excel、Word、PowerPoint、PDF、一太郎、画像ファイルを自動判定して**それっぽい**Markdownに変換するツール。

本ツールは[Devin](https://app.devin.ai)を利用して作成しています。

ちゃんとしたきゃ、[markitdown](https://github.com/microsoft/markitdown)や[docling](https://github.com/docling-project/docling)を使いましょう。

docling、markitdown との機能比較は [o2md_comparison.md](o2md_comparison.md) を参照してください。

## 概要

o2mdは、Microsoft Office文書（Excel、Word、PowerPoint）、PDF、一太郎文書（.jtd/.jtt）、および画像ファイル（JPEG/PNG/GIF/BMP/TIFF/WebP）を**それっぽい**Markdown形式に変換するPythonツールです。ファイルの種類を自動判定し、適切な変換エンジンを使用して処理します。
画像ファイルや画像ベースのPDF（スキャン文書等）は、OCR（Tesseract/manga-ocr/sarashina2.2-ocr）によりテキスト抽出を行います。
古い形式（.xls, .doc, .ppt）の変換、図形の画像処理と変換は**LibreOffice**に依存しています。LibreOfficeがない環境でもテキストのみの変換は正常に動作します。

### 主な特徴

- **統合インターフェース**: 1つのコマンドで全てのOffice文書とPDFを変換
- **自動ファイル判定**: ファイル拡張子に基づいて自動的に適切な変換方法を選択
- **新旧両形式対応**: `.xlsx`/`.xls`、`.docx`/`.doc`、`.pptx`/`.ppt`に対応
- **SVG/PNG出力対応**: 図形やグラフをSVG（デフォルト）またはPNG形式で出力
- **Excel変換** (x2md.py): 表、グラフ、図形を含むワークシートを変換
- **Word変換** (d2md.py): 見出し、表、画像、リストを含む文書を変換
- **PowerPoint変換** (p2md.py): スライド、図形、表、テキストを変換
- **PDF変換** (pdf2md.py): PDFを画像とテキストに変換（[Tesseract](https://github.com/tesseract-ocr/tesseract)/[manga-ocr](https://github.com/kha-white/manga-ocr)/[sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr)によるOCR対応）
- **一太郎変換** (jtd2md.py): 一太郎文書のOLE2バイナリを独自解析し、テキスト・テーブル・太字を変換
- **画像OCR変換** (img2md.py): 画像ファイルからOCRでテキスト抽出しMarkdownに変換（Tesseract/manga-ocr/sarashina対応）
- **画像処理**: 図形やグラフを自動的に画像として抽出・埋め込み
- **複雑な要素の処理**: 表と図形が混在するスライドは全体を画像化

※[sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr)すごいです。オススメです。GPU等あるなら是非www

## インストール

### 前提条件

- Python 3.10 以上
- [uv](https://docs.astral.sh/uv/) (推奨) または pip
- [LibreOffice](https://www.libreoffice.org/) (図形の画像処理と古い形式の変換に必要、オプショナル)

### 1. uv のインストール（未インストールの場合）

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 2. プロジェクトのセットアップ

```bash
# リポジトリをクローン
git clone https://github.com/matsu582/o2md.git
cd o2md

# 依存関係をインストール（uv sync で pyproject.toml から自動インストール）
uv sync

# 開発用依存関係も含める場合
uv sync --all-extras
```

### 3. LibreOffice のインストール（オプショナル）

古い形式（.xls, .doc, .ppt）の変換、図形の画像処理と変換に必要です。
インストールしない場合でも、テキストのみの変換は正常に動作します。

```bash
# macOS
brew install libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice

# Windows
# https://www.libreoffice.org/download/download/ からダウンロード
```

### LibreOfficeがない場合の動作

LibreOfficeがインストールされていない環境では、起動時に警告メッセージが表示され、以下のように縮退動作します。

| 機能 | LibreOfficeあり | LibreOfficeなし |
| --- | --- | --- |
| .docx / .xlsx / .pptx のテキスト変換 | ✔ | ✔ |
| .doc / .xls / .ppt の変換 | ✔ | ✖（エラー表示） |
| 図形・ベクター画像の変換 | ✔ | ✖（スキップ） |
| スライドの画像レンダリング | ✔ | ✖（スキップ） |
| チャートの画像変換 | ✔ | ✖（スキップ） |

> **ヒント**: 旧形式ファイル（.doc, .xls, .ppt）は、事前に新形式（.docx, .xlsx, .pptx）に変換しておくと、LibreOfficeなしでもテキスト変換が可能です。

## 使用方法

### 基本的な使用方法

```bash
# Excelファイルを変換
uv run python o2md.py input_files/data.xlsx

# Wordファイルを変換
uv run python o2md.py input_files/document.docx

# PowerPointファイルを変換
uv run python o2md.py input_files/presentation.pptx

# PDFファイルを変換
uv run python o2md.py input_files/document.pdf

# 一太郎ファイルを変換
uv run python o2md.py input_files/document.jtd

# 画像ファイルを変換（OCRでテキスト抽出）
uv run python o2md.py input_files/photo.jpg
```

### オプション

```bash
# 出力ディレクトリを指定
uv run python o2md.py input_files/data.xlsx -o custom_output

# Word文書で見出しテキストをリンクに使用
uv run python o2md.py input_files/document.docx --use-heading-text

# PNG形式で画像を出力（デフォルトはSVG）
uv run python o2md.py input_files/data.xlsx --format png

# PDF変換でOCRエンジンを指定（デフォルト: tesseract）
uv run python o2md.py input_files/document.pdf --ocr-engine tesseract
uv run python o2md.py input_files/document.pdf --ocr-engine manga-ocr
uv run python o2md.py input_files/document.pdf --ocr-engine sarashina

# tessdata_bestを使用する場合（高精度モード）
uv run python o2md.py input_files/document.pdf --tessdata-dir ~/tessdata_best

# テキスト抽出モード（.txtのみを出力）
uv run python o2md.py input_files/data.xlsx --text
uv run python o2md.py input_files/document.docx --text
uv run python o2md.py input_files/presentation.pptx --text
uv run python o2md.py input_files/document.pdf --text
uv run python o2md.py input_files/photo.jpg --text
```

### フォルダ一括変換

```bash
# フォルダ内の全対象ファイルを一括変換（フォルダ直下のみ）
uv run python o2md.py ./input_files/

# サブフォルダも再帰的に処理
uv run python o2md.py ./input_files/ -r

# 出力先を指定
uv run python o2md.py ./input_files/ -r -o output_all
```

フォルダ指定時の出力構造:
```
input_files/
  pdfs/a.pdf
  b.xlsx
↓
output/
  pdfs/a.md + images/
  b.md
```

### 古い形式のファイル

```bash
# 古い形式も自動的に新形式に変換してから処理
uv run python o2md.py input_files/old_file.xls
uv run python o2md.py input_files/old_doc.doc
uv run python o2md.py input_files/old_presentation.ppt
```

## コマンドラインオプション

| オプション           | 説明                                                    |
| -------------------- | ------------------------------------------------------- |
| `file`               | 変換するOfficeファイルまたはフォルダ（必須）                |
| `-o, --output-dir`   | 出力ディレクトリを指定（デフォルト: `./output`）        |
| `-r, --recursive`    | [フォルダ指定時] サブフォルダも再帰的に処理              |
| `--format`           | 画像出力形式を指定（`svg`または`png`、デフォルト: `svg`）|
| `--use-heading-text` | [Word専用] 章番号の代わりに見出しテキストをリンクに使用 |
| `--shape-metadata`   | [Word/Excel専用] 図形のメタデータを出力                 |
| `--ocr-engine`       | [PDF/画像] OCRエンジンを指定（`tesseract`/`manga-ocr`/`sarashina`、デフォルト: `tesseract`）|
| `--tessdata-dir`     | [PDF専用] tessdataディレクトリを指定（tessdata_best使用時）|
| `--docling`          | [PDF専用] doclingによる表検出を有効にする（罫線のない表も検出可能）|
| `--text`             | テキスト抽出モード（.txtのみを出力）                            |
| `-v, --verbose`      | 詳細なデバッグ出力を表示                                |
| `-h, --help`         | ヘルプメッセージを表示                                  |

## 対応ファイル形式

| ファイル種類 | 拡張子          | 変換エンジン | 主な機能                     |
| ------------ | --------------- | ------------ | ---------------------------- |
| Excel        | `.xlsx`, `.xls` | x2md.py      | 表、グラフ、図形、数式       |
| Word         | `.docx`, `.doc` | d2md.py      | 見出し、表、画像、リスト     |
| PowerPoint   | `.pptx`, `.ppt` | p2md.py      | スライド、図形、表、テキスト |
| PDF          | `.pdf`          | pdf2md.py    | 画像変換、テキスト抽出、OCR  |
| 一太郎       | `.jtd`, `.jtt`  | jtd2md.py    | テキスト、表、太字、見出し   |
| 画像         | `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.tif`, `.webp` | img2md.py | OCRテキスト抽出、画像埋め込み |

## 出力形式

変換後、以下のファイルが生成されます：

```
output/
├── [元のファイル名].md    # Markdownファイル
└── images/               # 画像フォルダ
    ├── [ファイル名]_image_001.svg  # デフォルトはSVG形式
    ├── [ファイル名]_image_002.svg
    └── ...
```

SVG形式はベクター形式のため、拡大しても品質が劣化しません。PNG形式が必要な場合は`--format png`オプションを使用してください。

## 変換機能の詳細

### Excel変換 (x2md.py)

- ワークシート内の表を Markdownテーブルに変換
- グラフを個別に画像として抽出（棒グラフ、折れ線グラフ、円グラフ、散布図に対応）
- グラフデータをMarkdownテーブルとして出力（画像の下にデータを配置）
- 図形（オートシェイプ、画像など）を画像化
- 数式の値を出力
- 複数シートの処理
- 図形クラスタリングによる分離レンダリング
- 罫線ベースのテーブル検出

### Word変換 (d2md.py)

- 見出しレベルを維持（`#`, `##`, `###` など）
- 段落とテキスト装飾（太字、斜体、下線、取り消し線、上付き/下付き文字）
- 箇条書きと番号付きリスト
- 表を Markdownテーブルに変換
- 埋め込み画像の抽出
- ハイパーリンクの保持
- 目次の自動生成
- 章参照のリンク変換（「第1章」→ `[第1章](#anchor)`）
- 数式変換（OMML → LaTeX）
- 図形・キャンバスの画像化
- チャートを文書内の元の位置に画像として出力（棒グラフ、折れ線グラフ、円グラフ、散布図に対応）
- チャートデータをMarkdownテーブルとして出力（画像の下にデータを配置）

### PowerPoint変換 (p2md.py)

- スライドごとに見出し設定
- テキストボックスの段落とリスト
- 表を Markdownテーブルに変換
- 図形群を1つの画像として出力
- 発表者ノートの抽出
- **複合スライド対応**: 表や図形が混在する場合、スライド全体を画像化してテキストを併記
- .pptファイル対応: LibreOfficeで自動変換

#### PowerPoint複合スライドの処理

スライドに以下の要素が混在する場合、スライド全体が画像化されます：

- テキストボックス + 図形
- 表 + 図形
- 複雑なレイアウト（視覚的な装飾を含む図形）

**視覚的装飾の判定基準**:
- 塗りつぶし（SOLID）がある
- 枠線の幅が0より大きい

これにより、吹き出しやカラフルな矩形などの装飾要素が適切に画像化されます。

### 一太郎変換 (jtd2md.py)

- OLE2 Compound Document形式のバイナリを独自パーサーで解析
- DocumentTextストリームからUTF-16BEテキストを抽出
- 罫線情報によるテーブル構造の検出とMarkdownテーブル出力
- フォントサイズによる見出し推定（本文より大きいフォント → `##`）
- TAG 0020（文字書式）の有無による太字検出（`**太字**`）
- Markdownビューワ対応の行末スペース付与
- 脚注テキストの抽出
- OLE2メタデータ（作成日時等）の出力

### 画像OCR変換 (img2md.py)

- 画像ファイル（JPEG/PNG/GIF/BMP/TIFF/WebP）からOCRでテキスト抽出
- 元画像をimagesフォルダにコピーし、Markdown内に画像リンクを埋め込み
- OCRエンジン選択: Tesseract（デフォルト）/ manga-ocr / sarashina2.2-ocr
- 日本語パス対応（cv2.imdecodeによる読み込み）
- `--text`オプション対応（.txtのみ出力）

### PDF変換 (pdf2md.py)

- 埋め込みテキスト・表の抽出
- スキャンページ（画像ベースPDF）は画像化してOCRでテキスト抽出
  - [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)（デフォルト）: 文書向けOCR、日本語+英語対応
  - [manga-ocr](https://github.com/kha-white/manga-ocr) + [comic-text-detector](https://github.com/dmMaze/comic-text-detector): マンガ/コミック向けOCR
  - [sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr): End-to-End VLMによる高精度OCR（GPU推奨）
- **[tessdata_best](https://github.com/tesseract-ocr/tessdata_best)対応**: `--tessdata-dir`オプションで高精度モデルを指定可能
- **[docling](https://github.com/docling-project/docling)表検出対応**: `--docling`オプションで罫線のない表も検出可能
  - [TableFormer](https://github.com/docling-project/docling)モデルを使用した高精度な表検出
  - スライドPDFや図形として描画された表に対応
  - 検出した表は`<details>`タグで囲んで出力
- 出力形式: ページごとの画像 + Markdownファイル

## 制限事項

### Excel (x2md.py)
- マクロは変換されません
- 複雑な条件付き書式は保持されません
- ピボットテーブルは静的なテーブルとして出力されます
- LibreOfficeの制限により図形の描画に差異が発生する場合があります。

### Word (d2md.py)
- 複雑なレイアウト（段組み、テキストボックス）は簡略化されます
- 脚注と文末注は通常のテキストとして処理されます
- LibreOfficeの制限により図形の描画に差異が発生する場合があります。
- コメントは変換されません

### PowerPoint (p2md.py)
- アニメーションは変換されません
- 埋め込み動画は変換されません（静止画のみ）
- スライドマスターのデザイン要素は反映されません
- LibreOfficeの制限により図形の描画に差異が発生する場合があります。

### PDF (pdf2md.py)
- 暗号化されたPDFは処理できません
- 複雑なレイアウトのPDFではテキスト抽出の精度が低下する場合があります
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)を使用するには事前にTesseractのインストールが必要です
  - macOS: `brew install tesseract tesseract-lang`
  - Ubuntu/Debian: `sudo apt-get install tesseract-ocr tesseract-ocr-jpn tesseract-ocr-eng`
- [tessdata_best](https://github.com/tesseract-ocr/tessdata_best)を使用する場合は別途ダウンロードが必要です（`--tessdata-dir`オプションで指定）
- [docling](https://github.com/docling-project/docling)表検出を使用するには追加のインストールが必要です
  - `uv pip install -e '.[docling]'`でdoclingをインストール
  - macOS ARM64の場合は`uv pip install rapidocr-torch`も必要
  - 処理時間: 1ページあたり約7-15秒（CPUのみの場合）
- [sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr)を使用するには追加のインストールが必要です
  - `uv pip install '.[sarashina]'`でインストール
  - GPU推奨（CUDA 8GB+ / Apple Silicon MPS 16GB+統合メモリ）
  - 初回実行時にモデル（約7.8GB）を自動ダウンロード
  - 画像→構造化Markdownを直接出力（テキスト検出不要）
  - 日本語縦書き・表・数式に対応

### 一太郎 (jtd2md.py)
- 一太郎 ver8以降のOLE2形式のみ対応（ver5-7の古い形式は非対応）
- 画像・図形の抽出は非対応
- 太字検出は段落単位（インラインの文字装飾は非対応）
- セル結合のある複雑な表はレイアウトが崩れる場合があります

### 画像OCR変換 (img2md.py)
- OCRの精度は画像の品質・解像度に依存します
- 手書き文字の認識精度は低い場合があります
- Tesseract OCRを使用するには事前にインストールが必要です
- sarashina2.2-ocrを使用するには`uv pip install '.[sarashina]'`が必要です（GPU推奨）

