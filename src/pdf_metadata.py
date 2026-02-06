"""Set PDF metadata (title, source) and get page count using pymupdf."""

import os
import shutil
import tempfile
from pathlib import Path


def set_pdf_metadata(
    pdf_path: str | Path,
    *,
    title: str | None = None,
    source: str | None = None,
) -> None:
    """Update PDF metadata. No-op if pymupdf unavailable or file missing."""
    try:
        import fitz  # pymupdf
    except ImportError:
        return
    path = Path(pdf_path)
    if not path.is_file():
        return
    tmp_path: str | None = None
    try:
        doc = fitz.open(path)
        meta = dict(doc.metadata) if doc.metadata else {}
        if title is not None:
            meta["title"] = title
        if source is not None:
            meta["producer"] = f"SlideToObsidian; source={source}"
        doc.set_metadata(meta)
        fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        doc.save(tmp_path)
        doc.close()
        doc = None
        shutil.move(tmp_path, path)
        tmp_path = None
    except Exception:
        pass
    finally:
        if tmp_path and Path(tmp_path).exists():
            try:
                Path(tmp_path).unlink()
            except Exception:
                pass


def get_pdf_page_count(pdf_path: str | Path) -> int:
    """Return the number of pages in the PDF. Returns 0 if unavailable."""
    try:
        import fitz  # pymupdf
    except ImportError:
        return 0
    path = Path(pdf_path)
    if not path.is_file():
        return 0
    try:
        doc = fitz.open(path)
        n = len(doc)
        doc.close()
        return n
    except Exception:
        return 0
