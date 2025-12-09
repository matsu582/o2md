# Office to Markdown Converter (o2md)

Excel、Word、PowerPointファイルを自動判定してMarkdownに変換するツール

## 概要

o2mdは、Microsoft Office文書（Excel、Word、PowerPoint）を**それっぽい**Markdown形式に変換するPythonツールです。ファイルの種類を自動判定し、適切な変換エンジンを使用して処理します。

### 主な特徴

- **統合インターフェース**: 1つのコマンドで全てのOffice文書を変換
- **自動ファイル判定**: ファイル拡張子に基づいて自動的に適切な変換方法を選択
- **新旧両形式対応**: `.xlsx`/`.xls`、`.docx`/`.doc`、`.pptx`/`.ppt`に対応
- **Excel変換** (x2md.py): 表、グラフ、図形を含むワークシートを変換
- **Word変換** (d2md.py): 見出し、表、画像、リストを含む文書を変換
- **PowerPoint変換** (p2md.py): スライド、図形、表、テキストを変換
- **画像処理**: 図形やグラフを自動的に画像として抽出・埋め込み
- **複雑な要素の処理**: 表と図形が混在するスライドは全体を画像化

## インストール

### 1. Pythonライブラリ

```bash
# pip を使用する場合
pip install openpyxl python-docx python-pptx Pillow PyMuPDF

# uv を使用する場合
uv pip install openpyxl python-docx python-pptx Pillow PyMuPDF
```

### 2. 外部ツール

#### LibreOffice
古い形式（.xls, .doc, .ppt）の変換、図形の画像処理と変換に必要

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
python o2md.py input_files/data.xlsx

# Wordファイルを変換
python o2md.py input_files/document.docx

# PowerPointファイルを変換
python o2md.py input_files/presentation.pptx
```

### オプション

```bash
# 出力ディレクトリを指定
python o2md.py input_files/data.xlsx -o custom_output

# Word文書で見出しテキストをリンクに使用
python o2md.py input_files/document.docx --use-heading-text
```

### 古い形式のファイル

```bash
# 古い形式も自動的に新形式に変換してから処理
python o2md.py input_files/old_file.xls
python o2md.py input_files/old_doc.doc
python o2md.py input_files/old_presentation.ppt
```

## コマンドラインオプション

| オプション           | 説明                                                    |
| -------------------- | ------------------------------------------------------- |
| `file`               | 変換するOfficeファイル（必須）                          |
| `-o, --output-dir`   | 出力ディレクトリを指定（デフォルト: `./output`）        |
| `--use-heading-text` | [Word専用] 章番号の代わりに見出しテキストをリンクに使用 |
| `-h, --help`         | ヘルプメッセージを表示                                  |

## 対応ファイル形式

| ファイル種類 | 拡張子          | 変換エンジン | 主な機能                     |
| ------------ | --------------- | ------------ | ---------------------------- |
| Excel        | `.xlsx`, `.xls` | x2md.py      | 表、グラフ、図形、数式       |
| Word         | `.docx`, `.doc` | d2md.py      | 見出し、表、画像、リスト     |
| PowerPoint   | `.pptx`, `.ppt` | p2md.py      | スライド、図形、表、テキスト |

## 出力形式

変換後、以下のファイルが生成されます：

```
output/
├── [元のファイル名].md    # Markdownファイル
└── images/               # 画像フォルダ
    ├── [ファイル名]_image_001.png
    ├── [ファイル名]_image_002.png
    └── ...
```

## 変換機能の詳細

### Excel変換 (x2md.py)

- ワークシート内の表を Markdownテーブルに変換
- グラフを画像として抽出
- 図形（オートシェイプ、画像など）を画像化
- 数式の値を出力
- 複数シートの処理

### Word変換 (d2md.py)

- 見出しレベルを維持（`#`, `##`, `###` など）
- 段落とテキスト装飾
- 箇条書きと番号付きリスト
- 表を Markdownテーブルに変換
- 埋め込み画像の抽出
- ハイパーリンクの保持

### PowerPoint変換 (p2md.py)

- スライドごとに見出し設定
- テキストボックスの段落とリスト
- 表を Markdownテーブルに変換
- 図形群を1つの画像として出力
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

## 使用例

```bash
# Excelファイルを変換
python o2md.py input_files/sales_data.xlsx

# Wordファイルを変換
python o2md.py input_files/manual.docx

# PowerPointファイルを変換
python o2md.py input_files/presentation.pptx
```

## 制限事項

### Excel (x2md.py)
- マクロは変換されません
- 複雑な条件付き書式は保持されません
- ピボットテーブルは静的なテーブルとして出力されます

### Word (d2md.py)
- 複雑なレイアウト（段組み、テキストボックス）は簡略化されます
- 脚注と文末注は通常のテキストとして処理されます
- コメントは変換されません

### PowerPoint (p2md.py)
- アニメーションは変換されません
- 埋め込み動画は変換されません（静止画のみ）
- 発表者ノートは変換されません
- スライドマスターのデザイン要素は反映されません

