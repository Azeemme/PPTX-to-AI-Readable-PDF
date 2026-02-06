# SlideToObsidian

High-performance Python CLI and GUI that batch-convert **PowerPoint files** into **AI-optimized Markdown + PDF pairs** for [Obsidian](https://obsidian.md/).

**Supported formats:** `.pptx`, `.ppt`, `.pot`, `.potx`, `.pps`, `.ppsx` (PDF and MD output use the same base name; speaker notes and rich body text are best for `.pptx`).

## Features

- **Dual output per file:** e.g. `Lecture.pptx` or `Lecture.ppt` → `Lecture.pdf` (visual) + `Lecture.md` (AI-ready note)
- **Parallel pipeline:** Uses `ProcessPoolExecutor` with **all CPU cores minus one**
- **Semantic extraction:** [markitdown](https://github.com/microsoft/markitdown) for body text (tables, hierarchies preserved)
- **Speaker notes:** Extracted with `python-pptx` and placed in YAML frontmatter
- **Obsidian embeds:** Each slide gets `![[Lecture.pdf#page=N]]` plus that slide’s text
- **Resilience:** Per-file try/except, progress bar (`tqdm`), LibreOffice check before run
- **PDF metadata:** Title and source set with pymupdf

## Requirements

- **Python 3.11+**
- **LibreOffice** (for headless PDF export). Install from [libreoffice.org](https://www.libreoffice.org/download) or set `LIBREOFFICE_PATH` to your `soffice` executable.

## Install

```bash
pip install -r requirements.txt
```

## Usage

### GUI (recommended)

Set the directory and click **Process** to convert all PowerPoint files in that folder (and subfolders). Output is written to an `out/` folder inside the selected directory.

```bash
python main.py --gui
```

### CLI

```bash
# Convert a single file → out/Lecture.pdf + out/Lecture.md
python main.py path/to/Lecture.pptx
python main.py path/to/Lecture.ppt

# Convert all PowerPoint files under a directory (output to out/, mirror structure by default)
python main.py path/to/slides/ -o out/

# Flat output (all files in one directory)
python main.py path/to/slides/ -o out/ --no-mirror
```

| Argument | Description |
|----------|-------------|
| `--gui` | Launch the GUI (directory picker + Process button) |
| `input` | File or directory; directories are scanned for `.pptx`, `.ppt`, `.pot`, `.potx`, `.pps`, `.ppsx` |
| `-o, --output` | Output directory (default: `out/`) |
| `--no-mirror` | Put all outputs in one directory instead of mirroring structure |

If LibreOffice is not found, the script prints a clear error and exits.

## Project layout

```
├── main.py              # CLI entry point (and --gui)
├── gui.py               # GUI (directory picker + Process button)
├── requirements.txt
├── src/
│   ├── __init__.py
│   ├── converter.py     # Single-file pipeline (PDF + MD)
│   ├── libreoffice.py   # LibreOffice check + headless conversion
│   ├── markdown_builder.py  # Frontmatter + per-slide embeds
│   ├── pdf_metadata.py  # pymupdf metadata
│   ├── pptx_utils.py    # python-pptx (speaker notes, slide count)
│   └── worker.py        # ProcessPool worker (reuses MarkItDown per process)
```

## Tech stack

| Component | Role |
|-----------|------|
| **markitdown** | Semantic body text (tables, structure) |
| **python-pptx** | Speaker notes, slide count, alt-text |
| **LibreOffice (headless)** | PPTX → PDF |
| **pymupdf** | PDF metadata (title, source) |
| **concurrent.futures.ProcessPoolExecutor** | Parallel workers (cores - 1) |
| **tqdm** | Progress bar |

## License

See [LICENSE](LICENSE).
