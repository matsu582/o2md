# Office to Markdown Converter (o2md)

Excel、Word、PowerPoint、PDFファイルを自動判定して**それっぽい**Markdownに変換するツール。

本ツールは[Devin](https://app.devin.ai)を利用して作成しています。

ちゃんとしたきゃ、[markitdown](https://github.com/microsoft/markitdown)や[docling](https://github.com/docling-project/docling)を使いましょう。

docling、markitdown との機能比較は [o2md_comparison.md](o2md_comparison.md) を参照してください。

## 概要

o2mdは、Microsoft Office文書（Excel、Word、PowerPoint）を**それっぽい**Markdown形式に変換するPythonツールです。ファイルの種類を自動判定し、適切な変換エンジンを使用して処理します。
古い形式（.xls, .doc, .ppt）の変換、図形の画像処理と変換は**LibreOffice**に依存しています。必ず**LibreOffice**をインストールしてください。

### 主な特徴

- **統合インターフェース**: 1つのコマンドで全てのOffice文書を変換
- **自動ファイル判定**: ファイル拡張子に基づいて自動的に適切な変換方法を選択
- **新旧両形式対応**: `.xlsx`/`.xls`、`.docx`/`.doc`、`.pptx`/`.ppt`に対応
- **SVG/PNG出力対応**: 図形やグラフをSVG（デフォルト）またはPNG形式で出力
- **Excel変換** (x2md.py): 表、グラフ、図形を含むワークシートを変換
- **Word変換** (d2md.py): 見出し、表、画像、リストを含む文書を変換
- **PowerPoint変換** (p2md.py): スライド、図形、表、テキストを変換
- **PDF変換** (pdf2md.py): PDFを画像とテキストに変換（OCRフォールバック対応）
- **画像処理**: 図形やグラフを自動的に画像として抽出・埋め込み
- **複雑な要素の処理**: 表と図形が混在するスライドは全体を画像化

## インストール

### 前提条件

- Python 3.9 以上
- [uv](https://docs.astral.sh/uv/) (推奨) または pip
- LibreOffice (図形の画像処理と古い形式の変換に必要)

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

### 3. LibreOffice のインストール

古い形式（.xls, .doc, .ppt）の変換、図形の画像処理と変換に必要です。

```bash
# macOS
brew install libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice

# Windows
# https://www.libreoffice.org/download/download/ からダウンロード
```

## 使用方法

### 基本的な使用方法

```bash
# Excelファイルを変換
uv run python o2md.py input_files/data.xlsx

# Wordファイルを変換
uv run python o2md.py input_files/document.docx

# PowerPointファイルを変換
uv run python o2md.py input_files/presentation.pptx

# PDFファイルを画像とテキストに変換
uv run python pdf2md.py /path/to/pdf/directory
```

### オプション

```bash
# 出力ディレクトリを指定
uv run python o2md.py input_files/data.xlsx -o custom_output

# Word文書で見出しテキストをリンクに使用
uv run python o2md.py input_files/document.docx --use-heading-text

# PNG形式で画像を出力（デフォルトはSVG）
uv run python o2md.py input_files/data.xlsx --format png
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
| `file`               | 変換するOfficeファイル（必須）                          |
| `-o, --output-dir`   | 出力ディレクトリを指定（デフォルト: `./output`）        |
| `--format`           | 画像出力形式を指定（`svg`または`png`、デフォルト: `svg`）|
| `--use-heading-text` | [Word専用] 章番号の代わりに見出しテキストをリンクに使用 |
| `--shape-metadata`   | [Word/Excel専用] 図形のメタデータを出力                 |
| `-v, --verbose`      | 詳細なデバッグ出力を表示                                |
| `-h, --help`         | ヘルプメッセージを表示                                  |

## 対応ファイル形式

| ファイル種類 | 拡張子          | 変換エンジン | 主な機能                     |
| ------------ | --------------- | ------------ | ---------------------------- |
| Excel        | `.xlsx`, `.xls` | x2md.py      | 表、グラフ、図形、数式       |
| Word         | `.docx`, `.doc` | d2md.py      | 見出し、表、画像、リスト     |
| PowerPoint   | `.pptx`, `.ppt` | p2md.py      | スライド、図形、表、テキスト |
| PDF          | `.pdf`          | pdf2md.py    | 画像変換、テキスト抽出、OCR  |

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

### PDF変換 (pdf2md.py)

- PDFの各ページを画像ファイル（PNG/JPEG）に変換
- 埋め込みテキストの抽出
- **OCRフォールバック**: テキストが抽出できない場合はOCRで読み取り
- 複数PDFファイルの一括処理
- 出力形式: ページごとの画像 + PDFごとのテキストファイル

#### PDF変換の出力構造

```
output/
├── document1/
│   ├── images/
│   │   ├── page_001.png
│   │   ├── page_002.png
│   │   └── ...
│   └── document1.txt
├── document2/
│   ├── images/
│   │   └── ...
│   └── document2.txt
```

#### OCR機能を使用する場合

Tesseract OCRのインストールが必要です：

```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr tesseract-ocr-jpn

# macOS
brew install tesseract tesseract-lang
```

#### PowerPoint複合スライドの処理

スライドに以下の要素が混在する場合、スライド全体が画像化されます：

- テキストボックス + 図形
- 表 + 図形
- 複雑なレイアウト（視覚的な装飾を含む図形）

**視覚的装飾の判定基準**:
- 塗りつぶし（SOLID）がある
- 枠線の幅が0より大きい

これにより、吹き出しやカラフルな矩形などの装飾要素が適切に画像化されます。

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
- OCR機能を使用するにはTesseract OCRのインストールが必要です
- 複雑なレイアウトのPDFではテキスト抽出の精度が低下する場合があります
- 暗号化されたPDFは処理できません

