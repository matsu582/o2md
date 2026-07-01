# Office to Markdown Converter (o2md)

A tool that auto-detects Excel, Word, PowerPoint, PDF, Ichitaro, and image files and converts them to **roughly adequate** Markdown.

This tool was built using [Devin](https://app.devin.ai).

For production-grade conversion, consider [markitdown](https://github.com/microsoft/markitdown) or [docling](https://github.com/docling-project/docling).

For a feature comparison with docling and markitdown, see [o2md_comparison.md](o2md_comparison.md).

[Japanese documentation (README_JP.md)](README_JP.md)

## Overview

o2md is a Python tool that converts Microsoft Office documents (Excel, Word, PowerPoint, MS Project), PDFs, Ichitaro documents (.jtd/.jtt/.jsw/.jaw/.jtw/.jbw/.juw/.jfw/.jvw), and image files (JPEG/PNG/GIF/BMP/TIFF/WebP) to **roughly adequate** Markdown. It auto-detects file types and selects the appropriate conversion engine.
o2md does not use trendy machine learning-based conversion; instead, it performs good old-fashioned logic-based conversion.
Image files and image-based PDFs (scanned documents, etc.) are processed with OCR (Tesseract/manga-ocr/sarashina2.2-ocr) for text extraction.
Legacy formats (.xls, .doc, .ppt) and shape/image rendering depend on **LibreOffice**. Text-only conversion works without LibreOffice.

### Key Features

- **Unified interface**: Convert all Office documents and PDFs with a single command
- **Auto file detection**: Automatically selects the appropriate conversion method based on file extension
- **Legacy format support**: Supports `.xlsx`/`.xls`, `.docx`/`.doc`, `.pptx`/`.ppt`
- **SVG/PNG output**: Export shapes and charts as SVG (default) or PNG
- **Excel** (`x2md`): Convert worksheets including tables, charts, and shapes
- **Word** (`d2md`): Convert documents with headings, tables, images, and lists
- **PowerPoint** (`p2md`): Convert slides with shapes, tables, and text
- **PDF** (`pdf2md`): Convert PDFs with text extraction and OCR ([Tesseract](https://github.com/tesseract-ocr/tesseract)/[manga-ocr](https://github.com/kha-white/manga-ocr)/[sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr))
- **Ichitaro** (`jtd2md`): Parse OLE2 binary format (ver8+) and legacy binary format (ver4-7) to extract text, tables, and bold text
- **MS Project** (`mpp2md`): Convert project files (.mpp/.mpt/.mpx) to Markdown tables with tasks, schedules, progress, and resource assignments. Uses mpxj via JPype (requires JDK 11+, optional dependency: `pip install o2md[mpp]`)
- **Image OCR** (`img2md`): Extract text from images via OCR and convert to Markdown (Tesseract/manga-ocr/sarashina)
- **Image extraction**: Automatically extract and embed shapes and charts as images
- **Complex element handling**: Slides with mixed tables and shapes are rendered as full-slide images
- **Search engine filter** (`o2md-filter`): Convert documents to plain text for search engine indexing (stdin support with auto file type detection)
- **Fess integration**: Integrate o2md-filter with [Fess](https://github.com/codelibs/fess) full-text search server and use as an MCP server for AI assistants (see [fess/README.md](fess/README.md))

*[sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr) is amazing. Highly recommended if you have a GPU!*

## Installation

### Prerequisites

- Python 3.10+
- [uv](https://docs.astral.sh/uv/) (recommended) or pip
- [LibreOffice](https://www.libreoffice.org/) (optional, required for shape rendering and legacy format conversion)

### Install from PyPI

```bash
# Basic install
pip install o2md

# With manga-ocr
pip install o2md[manga-ocr]

# With sarashina2.2-ocr
pip install o2md[sarashina]

# With docling (AI table detection)
pip install o2md[docling]

# With MS Project support (requires JDK 11+)
pip install o2md[mpp]

# All optional dependencies
pip install o2md[manga-ocr,sarashina,docling,mpp]
```

Using uv:

```bash
# Basic install
uv pip install o2md

# With manga-ocr
uv pip install o2md[manga-ocr]

# With sarashina2.2-ocr
uv pip install o2md[sarashina]

# With MS Project support (requires JDK 11+)
uv pip install o2md[mpp]

# All optional dependencies
uv pip install o2md[manga-ocr,sarashina,docling,mpp]
```

### Local Development Setup

```bash
# Install uv (if not installed)
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Clone the repository
git clone https://github.com/matsu582/o2md.git
cd o2md

# Install dependencies
uv sync

# Include manga-ocr
uv sync --extra manga-ocr

# Include sarashina2.2-ocr
uv sync --extra sarashina

# Include docling (AI table detection)
uv sync --extra docling

# Include MS Project support (requires JDK 11+)
uv sync --extra mpp

# Include all optional dependencies
uv sync --all-extras
```

### Install Tesseract OCR (Optional)

Used by default for OCR text extraction from PDFs and images.
Without Tesseract, OCR is unavailable (manga-ocr or sarashina2.2-ocr can be used as alternatives).

```bash
# macOS
brew install tesseract tesseract-lang

# Ubuntu/Debian
sudo apt-get install tesseract-ocr tesseract-ocr-jpn tesseract-ocr-eng

# Windows
# Download from https://github.com/UB-Mannheim/tesseract/wiki
```

### Install LibreOffice (Optional)

Required for legacy format (.xls, .doc, .ppt) conversion and shape/image rendering.
Text-only conversion works without LibreOffice.

```bash
# macOS
brew install libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice

# Windows
# Download from https://www.libreoffice.org/download/download/
```

### Behavior Without LibreOffice

Without LibreOffice, a warning is displayed at startup and the following degraded behavior applies:

| Feature | With LibreOffice | Without LibreOffice |
| --- | --- | --- |
| .docx / .xlsx / .pptx text conversion | OK | OK |
| .doc / .xls / .ppt conversion | OK | Error |
| Shape/vector image conversion | OK | Skipped |
| Slide image rendering | OK | Skipped |
| Chart image conversion | OK | Skipped |

> **Tip**: Legacy files (.doc, .xls, .ppt) can be pre-converted to newer formats (.docx, .xlsx, .pptx) for text conversion without LibreOffice.

## Usage

### Basic Usage

After installing from PyPI:

```bash
# Convert Excel file
o2md data.xlsx

# Convert Word file
o2md document.docx

# Convert PowerPoint file
o2md presentation.pptx

# Convert PDF file
o2md document.pdf

# Convert Ichitaro file
o2md document.jtd

# Convert MS Project file
o2md project.mpp

# Convert image file (OCR text extraction)
o2md photo.jpg
```

When running from a local clone (development):

```bash
uv run o2md input_files/data.xlsx
uv run o2md input_files/document.docx
uv run o2md input_files/presentation.pptx
uv run o2md input_files/document.pdf
uv run o2md input_files/document.jtd
uv run o2md input_files/photo.jpg
uv run o2md input_files/project.mpp
```

### Individual Commands

Each conversion engine can also be used as a standalone command:

```bash
# PyPI install
d2md document.docx
x2md data.xlsx
p2md presentation.pptx
pdf2md document.pdf
jtd2md document.jtd
img2md photo.jpg
mpp2md project.mpp

# Local clone
uv run d2md input_files/document.docx
uv run x2md input_files/data.xlsx
uv run p2md input_files/presentation.pptx
uv run pdf2md input_files/document.pdf
uv run jtd2md input_files/document.jtd
uv run img2md input_files/photo.jpg
uv run mpp2md input_files/project.mpp
```

### Plain Text Filter (`o2md-filter`)

Converts documents to plain text and outputs to stdout.

```bash
# File path input
o2md-filter document.xlsx
o2md-filter report.pdf
o2md-filter manual.docx

# stdin input (auto file type detection via magic bytes)
cat document.xlsx | o2md-filter
cat report.pdf | o2md-filter

# Specify OCR engine
o2md-filter scanned.pdf --ocr-engine manga-ocr
o2md-filter image.jpg --ocr-engine sarashina

# Search engine integration example
find /docs -name "*.xlsx" -o -name "*.pdf" | while read f; do
  o2md-filter "$f" | index-tool --source "$f"
done
```

| Option | Description |
| --- | --- |
| `file` | File to convert (omit to read from stdin) |
| `--ocr-engine` | OCR engine (`tesseract`/`manga-ocr`/`sarashina`, default: `tesseract`) |

### Options

```bash
# Specify output directory
o2md data.xlsx -o custom_output

# Use heading text for links in Word documents
o2md document.docx --use-heading-text

# Output images in PNG format (default: SVG)
o2md data.xlsx --format png

# Specify OCR engine for PDF/image conversion (default: tesseract)
o2md document.pdf --ocr-engine tesseract
o2md document.pdf --ocr-engine manga-ocr
o2md document.pdf --ocr-engine sarashina

# Use tessdata_best for high-accuracy OCR
o2md document.pdf --tessdata-dir ~/tessdata_best

# Text extraction mode (output .txt only)
o2md data.xlsx --text
o2md document.docx --text
o2md presentation.pptx --text
o2md document.pdf --text
o2md photo.jpg --text
```

> **Note**: When running from a local clone, prefix commands with `uv run` (e.g., `uv run o2md data.xlsx`).

### Batch Folder Conversion

```bash
# Convert all supported files in a folder (top-level only)
o2md ./input_files/

# Recursively process subfolders
o2md ./input_files/ -r

# Specify output directory
o2md ./input_files/ -r -o output_all
```

Output structure for folder input:
```
input_files/
  pdfs/a.pdf
  b.xlsx
->
output/
  pdfs/a.md + images/
  b.md
```

### Legacy Format Files

```bash
# Legacy formats are automatically converted to newer formats before processing
o2md old_file.xls
o2md old_doc.doc
o2md old_presentation.ppt
```

## Command-Line Options

| Option               | Description                                              |
| -------------------- | -------------------------------------------------------- |
| `file`               | Office file or folder to convert (required)              |
| `-o, --output-dir`   | Output directory (default: `./output`)                   |
| `-r, --recursive`    | [Folder mode] Recursively process subfolders             |
| `--format`           | Image output format (`svg` or `png`, default: `svg`)     |
| `--use-heading-text` | [Word only] Use heading text instead of chapter numbers for links |
| `--shape-metadata`   | [Word/Excel only] Output shape metadata                  |
| `--ocr-engine`       | [PDF/Image] OCR engine (`tesseract`/`manga-ocr`/`sarashina`, default: `tesseract`) |
| `--tessdata-dir`     | [PDF only] Specify tessdata directory (for tessdata_best) |
| `--docling`          | [PDF only] Enable docling table detection (detects borderless tables) |
| `--text`             | Text extraction mode (output .txt only)                  |
| `-v, --verbose`      | Show detailed debug output                               |
| `-h, --help`         | Show help message                                        |

## Supported File Formats

| File Type    | Extensions      | Command    | Key Features                 |
| ------------ | --------------- | ---------- | ---------------------------- |
| Excel        | `.xlsx`, `.xls` | `x2md`     | Tables, charts, shapes, formulas |
| Word         | `.docx`, `.doc` | `d2md`     | Headings, tables, images, lists |
| PowerPoint   | `.pptx`, `.ppt` | `p2md`     | Slides, shapes, tables, text |
| PDF          | `.pdf`          | `pdf2md`   | Image conversion, text extraction, OCR |
| Ichitaro     | `.jtd`, `.jtt`, `.jsw`, `.jaw`, `.jtw`, `.jbw`, `.juw`, `.jfw`, `.jvw` | `jtd2md`   | Text, tables, bold, headings |
| MS Project   | `.mpp`, `.mpt`, `.mpx` | `mpp2md`   | Tasks, schedules, progress, resources |
| Image        | `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.tif`, `.webp` | `img2md` | OCR text extraction, image embedding |

## Output Format

After conversion, the following files are generated:

```
output/
+-- [filename].md           # Markdown file
+-- images/                  # Image folder
    +-- [filename]_image_001.svg  # SVG by default
    +-- [filename]_image_002.svg
    +-- ...
```

SVG is a vector format that maintains quality at any zoom level. Use `--format png` if PNG is needed.

## Conversion Details

### Excel (`x2md`)

- Convert worksheet tables to Markdown tables
- Extract charts as individual images (bar, line, pie, scatter)
- Output chart data as Markdown tables below chart images
- Convert shapes (autoshapes, images, etc.) to images
- Output formula values
- Multi-sheet processing
- Shape clustering with isolated rendering
- Border-based table detection

### Word (`d2md`)

- Preserve heading levels (`#`, `##`, `###`, etc.)
- Paragraph and text decoration (bold, italic, underline, strikethrough, superscript/subscript)
- Bullet and numbered lists
- Convert tables to Markdown tables
- Extract embedded images
- Preserve hyperlinks
- Auto-generate table of contents
- Chapter reference link conversion
- Equation conversion (OMML -> LaTeX)
- Shape and canvas image rendering
- Charts rendered as images at their original position (bar, line, pie, scatter)
- Chart data output as Markdown tables below chart images

### PowerPoint (`p2md`)

- Per-slide heading generation
- Text box paragraphs and lists
- Convert tables to Markdown tables
- Render shape groups as single images
- Extract speaker notes
- **Complex slide handling**: Slides with mixed tables and shapes are rendered as full-slide images with text alongside
- .ppt file support: Auto-conversion via LibreOffice

#### Complex Slide Handling

Slides are rendered as full images when they contain mixed elements:

- Text boxes + shapes
- Tables + shapes
- Complex layouts (shapes with visual decorations)

**Visual decoration criteria**:
- Has solid fill
- Has border width > 0

This ensures callouts, colored rectangles, and other decorative elements are properly rendered as images.

### Ichitaro (`jtd2md`)

- Parse OLE2 Compound Document binary format
- Extract UTF-16BE text from DocumentText stream
- Detect table structure from border information and output as Markdown tables
- Estimate headings from font size (larger than body text -> `##`)
- Detect bold text via TAG 0020 (character formatting) (`**bold**`)
- Add trailing spaces for Markdown viewer compatibility
- Extract footnote text
- Output OLE2 metadata (creation date, etc.)

### MS Project (`mpp2md`)

- Convert MS Project files (.mpp/.mpt/.mpx) to Markdown
- Output task list as a table with name, start date, end date, duration, progress, and assigned resources
- Support task hierarchy with indent-based nesting (summary tasks in bold)
- Resource list table output
- Project metadata (title, author, start/end dates)
- Uses mpxj library via JPype (JDK 11+ required)
- Optional dependency: `pip install o2md[mpp]`
- JARs are auto-downloaded from Maven Central on first use
- MSPDI XML (.xml) format also supported via `mpp2md` command directly

### Image OCR (`img2md`)

- Extract text from images (JPEG/PNG/GIF/BMP/TIFF/WebP) via OCR
- Copy original image to images folder and embed image link in Markdown
- OCR engine selection: Tesseract (default) / manga-ocr / sarashina2.2-ocr
- Japanese path support (via cv2.imdecode)
- `--text` option support (output .txt only)

### PDF (`pdf2md`)

- Extract embedded text and tables
- Scan pages (image-based PDFs) are rendered as images with OCR text extraction
  - [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (default): Document OCR, Japanese + English
  - [manga-ocr](https://github.com/kha-white/manga-ocr) + [comic-text-detector](https://github.com/dmMaze/comic-text-detector): Manga/comic OCR
  - [sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr): End-to-End VLM high-accuracy OCR (GPU recommended)
- **[tessdata_best](https://github.com/tesseract-ocr/tessdata_best) support**: Specify high-accuracy models via `--tessdata-dir`
- **[docling](https://github.com/docling-project/docling) table detection**: `--docling` option detects borderless tables
  - High-accuracy table detection using [TableFormer](https://github.com/docling-project/docling) model
  - Supports slide PDFs and tables drawn as shapes
  - Detected tables are wrapped in `<details>` tags
- Output: Per-page images + Markdown file

## Limitations

### Excel (`x2md`)
- Macros are not converted
- Complex conditional formatting is not preserved
- Pivot tables are output as static tables
- Shape rendering may differ due to LibreOffice limitations

### Word (`d2md`)
- Complex layouts (columns, text boxes) are simplified
- Footnotes and endnotes are processed as regular text
- Shape rendering may differ due to LibreOffice limitations
- Comments are not converted

### PowerPoint (`p2md`)
- Animations are not converted
- Embedded videos are not converted (still images only)
- Slide master design elements are not reflected
- Shape rendering may differ due to LibreOffice limitations

### PDF (`pdf2md`)
- Encrypted PDFs cannot be processed
- Text extraction accuracy may decrease for complex PDF layouts
- [tessdata_best](https://github.com/tesseract-ocr/tessdata_best) requires separate download (`--tessdata-dir` option)
- [docling](https://github.com/docling-project/docling) table detection requires additional installation
  - Install with `pip install o2md[docling]` or `uv sync --extra docling`
  - Processing time: ~7-15 seconds per page (CPU only)
- [manga-ocr](https://github.com/kha-white/manga-ocr) requires additional installation
  - Install with `pip install o2md[manga-ocr]` or `uv sync --extra manga-ocr`
- [sarashina2.2-ocr](https://huggingface.co/sbintuitions/sarashina2.2-ocr) requires additional installation
  - Install with `pip install o2md[sarashina]` or `uv sync --extra sarashina`
  - GPU recommended (CUDA 8GB+ / Apple Silicon MPS 16GB+ unified memory)
  - Model (~7.8GB) is auto-downloaded on first run
  - Direct image-to-structured-Markdown output (no text detection needed)
  - Supports Japanese vertical text, tables, and formulas

### Ichitaro (`jtd2md`)
- OLE2 format (Ichitaro ver8+) and legacy binary format (ver4-7) are supported
- Legacy format (ver4-6) text extraction has not been tested with real files due to their scarcity; only synthetic unit tests have been performed
- Image/shape extraction is not supported
- Bold detection is per-paragraph (inline character decoration is not supported)
- Complex tables with merged cells may have layout issues

### MS Project (`mpp2md`)
- Requires JDK/JRE 11+ and jpype1: `pip install o2md[mpp]`
- MPXJ JARs (~30MB) are auto-downloaded from Maven Central on first use
- Network access required for initial JAR download

### Image OCR (`img2md`)
- OCR accuracy depends on image quality and resolution
- Handwritten text recognition accuracy may be low
- Tesseract OCR requires prior system installation
- manga-ocr requires `pip install o2md[manga-ocr]` or `uv sync --extra manga-ocr`
- sarashina2.2-ocr requires `pip install o2md[sarashina]` or `uv sync --extra sarashina` (GPU recommended)

## Project Structure

```
o2md/
+-- pyproject.toml          # Project configuration and dependencies
+-- o2md/                   # Main package
|   +-- __init__.py
|   +-- o2md.py             # Unified CLI (auto file detection and conversion)
|   +-- x2md.py             # Excel conversion engine
|   +-- d2md.py             # Word conversion engine
|   +-- p2md.py             # PowerPoint conversion engine
|   +-- pdf2md.py           # PDF conversion engine
|   +-- jtd2md.py           # Ichitaro conversion engine
|   +-- jtd2md_legacy.py    # Ichitaro legacy (ver4-7) conversion engine
|   +-- mpp2md.py           # MS Project conversion engine
|   +-- jar_manager.py      # MPXJ JAR auto-download manager
|   +-- img2md.py           # Image OCR conversion engine
|   +-- filter.py           # Plain text filter for search engines
|   +-- utils.py            # Shared utilities
|   +-- omml_converter/     # OMML to LaTeX conversion
+-- tests/                  # Test suite
+-- input_files/            # Test input files
```
