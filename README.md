# txtFinder
**txtFinder** is a Python tool designed to search filenames from a structured JSON list across document collections and highlight every match directly in the source files. Built with forensic workflows in mind, it combines exact, fuzzy, stem-based, and regex search modes with automatic PDF report generation, hash list comparison, and file export — all through an interactive terminal menu (TUI).

## Why txtFinder?!
This tool was created out of the need to process over 1,500 non-searchable A4 PDF documents and check whether they contain references to one or more of 500+ filenames stored in a e01 Image.

The initial idea was to use OCR for text recognition. However, over time, the tool gradually grew beyond that original scope. At this point, I want to acknowledge that I am fully aware the current structure—especially having everything in a single file and lacking proper modularization—is far from ideal. It should have been planned more carefully instead of being extended “just in time.” That said, it works well for my purposes.

If you find this tool useful, I’m glad to hear it. Feel free to improve upon it and make up for my shortcomings.

> [!NOTE]
> txtFinder searches for filenames, not file contents. The filenames list (`filenames.json`) must be generated first and serves as the search pattern source for all file searches.

> [!WARNING]
> Please note that this script is currently under development, and I cannot provide a 100% guarantee that it operates in a forensically sound manner. It is tailored to meet specific needs at this stage. Use it with caution, especially in environments where forensic integrity is critical.

## Table of Contents
- [Features](#features)
- [Supported file types](#supported-file-types)
- [Requirements](#requirements)
- [Installation](#installation)
  - [Tesseract OCR (optional)](#tesseract-ocr-optional)
- [Usage](#usage)
  - [Typical workflow](#typical-workflow)
  - [Forensic hash comparison workflow](#forensic-hash-comparison-workflow)
  - [Search modes](#search-modes)
  - [Menu structure](#menu-structure)
- [Configuration](#configuration)
- [Output files](#output-files)
- [Project structure](#project-structure)
- [PyInstaller](#pyinstaller)
- [License](#license)
- [Contributing](#contributing)
- [Disclaimer](#disclaimer)

## Features
- **Interactive menu** — six menu items covering list generation, JSON search, file search, history, and settings
- **4 search modes** — exact substring, fuzzy (configurable threshold), word stems, and regular expressions
- **6 document types** — PDF (text-based and image/OCR), JPG/JPEG, DOCX, XLSX, PPTX, ODT, and TXT
- **OCR support** — reads text from JPG images and image-based PDFs via Tesseract; configurable language, PSM, DPI, and confidence threshold
- **Fuzzy matching** — sliding-window algorithm catches typos and OCR artefacts; default similarity threshold 80 %
- **Stems mode** — searches word stems to match inflected forms; configurable minimum stem length
- **Filelist comparison** — compare a plain-text list of filenames against `filenames.json`; generates a PDF report and an optional hitlist (`.txt`)
- **Hash list matching** — compare external SHA-256/MD5 hash lists against `filenames.json`; supports NSRL RDS CSV format; recursive `hashes/` directory scan
- **File export** — copy all hash-matched files into a structured `export/` directory with associated reports
- **Highlight & mark** — matches are annotated directly in output copies (PDF highlight annotations, `[[FOUND]]` markers in TXT, copied files for JPG and Office formats)
- **Per-file PDF reports** — match breakdown with file metadata, hash verification, context snippet, and a forensic disclaimer
- **CSV summary** — semicolon-delimited, timestamped export of all search results (UTF-8 BOM for broad tool compatibility)
- **Settings editor** — runtime editing of all parameters; changes are saved to `txtfinder_config.json` and loaded automatically on startup
- **Disk text cache** — extracted PDF and OCR text is persisted to `.txtfinder_cache/` with automatic 90-day cleanup
- **In-memory text cache** — extracted text is reused within a search run to avoid double-processing (FIFO eviction, 200-entry limit)
- **Search profiles** — extra search words can be saved as reusable profiles and loaded before each run

## Supported file types
| Type | Search | Output | OCR |
|------|--------|--------|-----|
| PDF (text-based) | ✓ | highlight annotation in `*_checked.pdf` | — |
| PDF (image-based) | ✓ | highlight annotation in `*_checked.pdf` | ✓ |
| JPG / JPEG | ✓ | copy as `*_found.jpg` | ✓ |
| DOCX | ✓ | copy as `*_checked.docx` | — |
| XLSX | ✓ | copy as `*_checked.xlsx` | — |
| PPTX | ✓ | copy as `*_checked.pptx` | — |
| ODT | ✓ | copy as `*_checked.odt` | — |
| TXT | ✓ | `[[FOUND]]` marker in `*_checked.txt` | — |

## Requirements
- **Python 3.10 or higher** (`match/case` syntax required)
- Tesseract OCR (optional — required only for JPG search and image-based PDF search)

Python dependencies:

```
PyMuPDF>=1.24.0,<2.0.0
typer>=0.12.0,<1.0.0
rapidfuzz>=3.0.0,<4.0.0
orjson>=3.9.0,<4.0.0
pytesseract>=0.3.10,<1.0.0
Pillow>=10.0.0,<12.0.0
python-docx>=1.0.0,<2.0.0
openpyxl>=3.1.0,<4.0.0
python-pptx>=1.0.0,<2.0.0
odfpy>=1.4.1,<2.0.0
```

## Installation
1. **Clone the repository**

```bash
git clone https://github.com/ot2i7ba/txtFinder.git
cd txtfinder
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

3. **Run the application**

```bash
python txtfinder.py
```

### Tesseract OCR (optional)
Tesseract is only required for searching JPG files and image-based PDFs. All other features work without it.

**Linux:**
```bash
sudo apt install tesseract-ocr tesseract-ocr-eng tesseract-ocr-deu
```

**macOS:**
```bash
brew install tesseract tesseract-lang
```

**Windows:**
Download and install from [UB Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and add Tesseract to your system `PATH`.

Alternatively, set the path directly in `txtfinder.py`:
```python
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

txtFinder automatically tries four common Windows installation paths before falling back to `PATH`.

## Usage
```bash
python txtfinder.py
```

The interactive menu opens:

```
╔══════════════════════════════════════════════════════╗
║  txtFinder v1.0 by ot2i7ba                           ║
║  Search · Highlight · Report                         ║
╚══════════════════════════════════════════════════════╝
  Dir: C:\cases\project01   OCR: ON   JSON: 12,847 entries (2026-03-18)

  1  Generate filenames list
  2  Generate hashlist file
  3  Search in filenames.json
  4  Search in Files
  5  Show search history
  6  Settings

  q  Quit

  Enter choice:
```

Navigation: `b` returns one level (submenu → parent menu), `q` quits from anywhere.

### Typical workflow
1. Place source files in `./input/` or any accessible directory
2. Run **1** → Generate filenames list — select hash mode (SHA-256 + MD5, SHA-256, MD5, no hashes, or stems only)
3. Run **4** → Search in Files — select the search directory, then choose a file type (PDF / JPG / Docs / TXT / all)
4. Review the generated `*_checked` output files and `*_report.pdf` per-file reports
5. Open `search_results.csv` for a full overview of all results

### Forensic hash comparison workflow
1. Place external hash list files (`.txt`, `.csv`) in the `hashes/` directory — subdirectories are scanned recursively
2. Run **3 → 5 (Match Hash List)** — txtFinder compares all hashes in `hashes/` against `filenames.json`
3. A PDF report and a CSV are generated; matched files can be exported immediately to `export/`
4. To re-export from a previous run: **3 → 6 (Export Matched Files)**

### Search modes
| Mode | Description |
|------|-------------|
| Exact | Fast case-insensitive substring matching (default) |
| Fuzzy | Catches typos and OCR artefacts; configurable similarity threshold (default 80 %) |
| Regex | Custom regular expressions |
| Stems | Matches word stems to catch inflected forms; configurable minimum stem length |

Additional search words can be entered before each run and saved as reusable search profiles.

### Menu structure
```
1  Generate filenames list
     1  Full list + SHA-256 & MD5 hashes
     2  Full list + SHA-256 only
     3  Full list + MD5 only
     4  Full list, no hashes
     5  Stems only
     (deduplication choice follows)

2  Generate hashlist file
     1  MD5
     2  SHA-256

3  Search in filenames.json
     1  Search by filename        (case-insensitive substring)
     2  Search for filelist       (compare a .txt file against filenames.json)
     3  Search by keyword         (filename + filepath)
     4  Search by hash            (MD5 or SHA-256 lookup → PDF report)
     5  Match hash list           (forensic comparison against hashes/ directory)
     6  Export matched files      (re-export from a previous *_hashlists.csv)

4  Search in Files
     (first: select search directory — input/ / custom path / CWD)
     1  Search in PDF
     2  Search in JPG
     3  Search in Docs (DOCX, XLSX, PPTX, ODT)
     4  Search in TXT
     a  Search in all file types

5  Show search history

6  Settings
```

## Configuration
All settings can be changed at the top of `txtfinder.py` or at runtime via **menu 6 (Settings)**. Runtime changes are saved to `txtfinder_config.json`, which is loaded automatically on startup.

| Constant | Default | Description |
|----------|---------|-------------|
| `HIGHLIGHT_COLOR` | `(1, 0, 0)` | PDF highlight color as RGB floats (0.0–1.0) |
| `TXT_MARKER` | `[[FOUND]]` | Marker text inserted in TXT file matches |
| `TXT_MARKER_POSITION` | `before` | Marker placement: `before` or `after` the match |
| `OUTPUT_DIR` | `""` | Output directory (empty = same directory as the source file) |
| `REPORT_ON_MATCH_ONLY` | `True` | Skip PDF report generation for files with zero matches |
| `OCR_MIN_CONFIDENCE` | `30` | Minimum Tesseract confidence score (0–100) |
| `OCR_LANGUAGES` | `eng+deu` | Tesseract language codes (e.g. `eng+deu+fra`) |
| `OCR_PSM` | `6` | Tesseract Page Segmentation Mode (3 = auto, 6 = block, 11 = sparse) |
| `OCR_DPI` | `200` | Resolution for rendering PDF pages before OCR (100–600) |
| `FUZZY_THRESHOLD` | `0.80` | Minimum similarity ratio for fuzzy matching (0.0–1.0) |
| `STEMS_MIN_LENGTH` | `5` | Minimum word stem length (0 = no limit) |
| `MAX_SEARCH_PATTERNS` | `500` | Maximum unique search patterns per run |
| `MAX_WORKERS` | `min(cpu, 8)` | Parallel worker threads (override via `TXTFINDER_WORKERS` env var) |
| `OUTPUT_SUFFIX` | `_checked` | Suffix appended to highlighted output files |
| `REPORT_SUFFIX` | `_report` | Suffix appended to per-file PDF reports |
| `HASHES_SUFFIX` | `_hashes` | Suffix appended to hash lookup report PDFs |
| `HASHLIST_REPORT_SUFFIX` | `_hashlists` | Suffix appended to hash list comparison reports |
| `FILELIST_REPORT_SUFFIX` | `_filelist` | Suffix appended to file list comparison reports |
| `HITLIST_SUFFIX` | `_hitlist` | Suffix appended to filelist hit result `.txt` files |
| `HASHLISTS_DIR` | `hashes` | Directory scanned recursively for external hash list files |
| `EXPORT_DIR` | `export` | Root directory for hash match file exports |
| `TESSERACT_PATH` | `None` | Explicit path to `tesseract.exe` on Windows (`None` = auto-detect) |
| `MEDIA_EXTENSIONS` | *(see code)* | File extensions included in media-only list generation |
| `DOCUMENT_EXTENSIONS` | *(see code)* | File extensions included in document-only list generation |

## Output files
| File / Directory | Description |
|------------------|-------------|
| `filenames.json` | Generated file list with paths, sizes, SHA-256 and MD5 hashes |
| `filenames_meta.json` | Companion metadata: scan root, hash mode, entry count, timestamp |
| `filenames_md5.txt` | MD5 hash list (one hash per line) |
| `filenames_sha256.txt` | SHA-256 hash list (one hash per line) |
| `*_checked.*` | Output files with highlighted or marked matches |
| `*_found.jpg` | JPG copies where OCR found a match |
| `*_report.pdf` | Per-file PDF report with match breakdown and forensic disclaimer |
| `*_hashes.pdf` | Hash lookup report (single hash or batch) |
| `*_hashlists.pdf` | Hash list comparison report |
| `*_hashlists.csv` | Hash list comparison CSV (UTF-8 BOM, semicolons) |
| `*_filelist.pdf` | File list comparison report |
| `*_hitlist.txt` | Matched filenames from a file list comparison (one per line) |
| `search_results.csv` | Timestamped CSV summary of all file search results |
| `search_history.json` | Automatic log of all search runs |
| `txtfinder_config.json` | Saved runtime settings (created via menu 6, auto-loaded on startup) |
| `txtfinder.log` | Rotating application log (max 1 MB, 3 backups) |
| `hashes/` | Place external hash list files here (subdirectories are scanned recursively) |
| `export/` | Exported files from hash list comparison runs |
| `.txtfinder_cache/` | Disk cache for extracted PDF and OCR text (90-day auto-cleanup) |

## Project structure
```
txtfinder/
├── input/                    # Default search directory for source files
├── hashes/                   # External hash list files (recursive scan)
├── export/                   # Hash match exports
├── search_profiles/          # Saved search profiles (JSON)
├── filenames.json            # Generated file list (auto-created)
├── filenames_meta.json       # Companion metadata (auto-created)
├── txtfinder_config.json     # Saved runtime settings (optional, auto-loaded)
├── search_history.json       # Search run history (auto-generated)
├── requirements.txt          # Python dependencies
├── txtfinder.log             # Application log (auto-generated)
└── txtfinder.py              # Main application
```

## PyInstaller
To create a single `txtFinder.exe` that runs on Windows without a Python installation:

**Install PyInstaller:**

```bash
pip install pyinstaller
```

You also need an icon file `txtfinder.ico` (Windows ICO format; recommended sizes: 16×16, 32×32, 48×48, 256×256).

**Windows build command:**

```bash
pyinstaller --onefile --console --name txtFinder --icon txtfinder.ico --noupx --noconfirm --clean --collect-all fitz --collect-all pymupdf --collect-all docx --collect-all pptx --collect-all odf --collect-all openpyxl --collect-all rapidfuzz --collect-all orjson --hidden-import pytesseract --hidden-import PIL._imaging --hidden-import typer txtfinder.py
```

The resulting executable is placed in `dist\txtFinder.exe`.

> [!IMPORTANT]
> Tesseract OCR is **not** bundled into the EXE and must be installed separately on each target machine. Download from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) and add it to the system `PATH`. Without Tesseract, the application runs normally — only JPG search and OCR-based PDF search are disabled, with a clear message shown.

## License
This project is licensed under the **[MIT license](https://github.com/ot2i7ba/txtFinder/blob/main/LICENSE)**, providing users with flexibility and freedom to use and modify the software according to their needs.

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request for review.

## Disclaimer
This project is provided without warranties of any kind. Results should always be verified manually. txtFinder is designed to support investigative workflows, not to replace them. Users are solely responsible for the interpretation and use of any output generated by this tool.

## Conclusion
This script has been tailored to fit my personal specific needs, and while it may seem simple, it has significant impact on my digital investigation workflows. This tool grew out of a concrete need in digital investigation work — hunting for specific filenames across thousands of documents is tedious by hand and error-prone under time pressure. txtFinder automates that process while keeping the analyst in control: no file is ever modified, every match is documented, and every report includes a clear disclaimer about the limits of automated search. Greetings to my dear colleagues who avoid scripts like the plague and think that consoles and Bash are some sort of dark magic – the [compiled](https://github.com/ot2i7ba/txtFinder/releases) version will spare you the console kung-fu and hopefully be a helpful tool for you as well. For everyone else: clone, install, run. 😉
