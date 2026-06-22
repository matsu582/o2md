# docling / markitdown / o2md 機能比較

## Word 変換機能

| 機能 | docling | markitdown | o2md |
|------|---------|------------|------|
| 基盤ライブラリ | [python-docx](https://github.com/python-openxml/python-docx) + lxml | mammoth (HTML経由) | [python-docx](https://github.com/python-openxml/python-docx) |
| 段落・テキスト | 対応 | 対応 | 対応 |
| 見出し | 対応 | 対応 | 対応 |
| 箇条書き/番号付きリスト | 対応 | 対応 | 対応 |
| 表 | 対応 | 対応 | 対応 |
| 画像抽出 | 対応 | 対応 | 対応 |
| 数式 (OMML→LaTeX) | 対応 | 対応 | 対応 |
| 図形処理 | [LibreOffice](https://www.libreoffice.org/)→PDF→PNG | なし | [LibreOffice](https://www.libreoffice.org/)→PDF→PNG/SVG |
| チャート | なし | なし | 対応 |
| 書式 (太字/斜体等) | 対応 | 対応 | 対応 |
| ヘッダー/フッター | 対応 | なし | なし |
| 目次生成 | なし | なし | 対応 |
| 章番号リンク変換 | なし | なし | 対応 |

## Excel 変換機能

| 機能 | docling | markitdown | o2md |
|------|---------|------------|------|
| 基盤ライブラリ | [openpyxl](https://openpyxl.readthedocs.io/) | pandas + [openpyxl](https://openpyxl.readthedocs.io/) | [openpyxl](https://openpyxl.readthedocs.io/) |
| テーブル検出 | データ領域自動検出 | シート全体をDataFrame | 罫線ベース検出 |
| セル結合 | 対応 | pandas依存 | 対応 |
| 複数シート | 対応 | 対応 | 対応 |
| 画像抽出 | 対応 | なし | 対応 |
| 図形処理 | なし | なし | [LibreOffice](https://www.libreoffice.org/)→PNG/SVG |
| チャート | なし | なし | 対応 |
| 書式 (太字/斜体等) | なし | なし | 対応 |
| 離散データ領域検出 | なし | なし | 対応 |
| 図形クラスタリング | なし | なし | 対応 |

## PowerPoint 変換機能

| 機能 | docling | markitdown | o2md |
|------|---------|------------|------|
| 基盤ライブラリ | [python-pptx](https://github.com/scanny/python-pptx) | [python-pptx](https://github.com/scanny/python-pptx) | [python-pptx](https://github.com/scanny/python-pptx) |
| スライドテキスト | 対応 | 対応 | 対応 |
| 表 | 対応 | 対応 | 対応 |
| 画像抽出 | 対応 | 対応 | 対応 |
| チャート | なし | 対応 | 対応 |
| 図形処理 | なし | なし | [LibreOffice](https://www.libreoffice.org/)→PNG/SVG |
| スライドノート | 対応 | 対応 | 対応 |
| 書式 (太字/斜体等) | 対応 | 対応 | 対応 |
| スライド画像化 | なし | なし | 対応 |
| 座標順ソート | なし | 対応 | 対応 |

## PDF 変換機能

| 機能 | docling | markitdown | o2md |
|------|---------|------------|------|
| 基盤ライブラリ | pypdfium2 + DocLayNet | pdfminer | [PyMuPDF](https://pymupdf.readthedocs.io/) (fitz) |
| テキスト抽出 | 対応 | 対応 | 対応 |
| 段組みレイアウト | 対応 | なし | 対応 |
| 表検出 | 対応 (TableFormer) | なし | 対応 (テキスト位置ベース + [docling](https://github.com/docling-project/docling) TableFormer) |
| 図抽出 | 対応 | なし | 対応 (SVG/PNG) |
| OCR | 対応 (EasyOCR/Tesseract) | なし | 対応 ([Tesseract](https://github.com/tesseract-ocr/tesseract)/[manga-ocr](https://github.com/kha-white/manga-ocr)/[sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr)選択可) |
| 書式 (太字/斜体等) | 対応 | なし | 対応 |
| 打消し線/上付き/下付き | なし | なし | 対応 |
| 脚注/注釈変換 | なし | なし | 対応 ([^N]形式) |
| ヘッダー/フッター除外 | 対応 | なし | 対応 |
| 番号付き箇条書き | 対応 | なし | 対応 |
| 図形内テキスト抽出 | なし | なし | 対応 |
| URL自動修復 | なし | なし | 対応 |
| スライド文書の図形クラスタリング | なし | なし | 対応（本文テキスト除外、フッタ除外） |

## 出力形式

| 項目 | docling | markitdown | o2md |
|------|---------|------------|------|
| Markdown | 対応 | 対応 | 対応 |
| 画像形式 | PNG | PNG | PNG/SVG選択可 |
| 中間表現 | DoclingDocument | HTML | なし（直接変換） |

## 参考リンク

- [docling](https://github.com/docling-project/docling) - IBM製ドキュメント変換ライブラリ
- [markitdown](https://github.com/microsoft/markitdown) - Microsoft製Markdown変換ツール
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) - オープンソースOCRエンジン
- [manga-ocr](https://github.com/kha-white/manga-ocr) - 日本語マンガ向けOCR
- [sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr) - 日本語Vision-Language OCRモデル
- [PyMuPDF](https://pymupdf.readthedocs.io/) - PDF処理ライブラリ
- [LibreOffice](https://www.libreoffice.org/) - オフィス文書変換に使用
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel (.xlsx) 読み書きライブラリ
- [python-docx](https://github.com/python-openxml/python-docx) - Word (.docx) 読み書きライブラリ
- [python-pptx](https://github.com/scanny/python-pptx) - PowerPoint (.pptx) 読み書きライブラリ
