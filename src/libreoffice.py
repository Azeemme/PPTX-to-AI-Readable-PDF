"""LibreOffice headless detection and PPTX → PDF conversion."""

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def _find_libreoffice() -> str | None:
    """Return path to soffice executable, or None if not found."""
    if sys.platform == "win32":
        names = ["soffice.exe", "soffice.com"]
        # Common install paths on Windows
        candidates = [
            os.environ.get("LIBREOFFICE_PATH"),
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    else:
        names = ["soffice", "libreoffice"]
        candidates = [os.environ.get("LIBREOFFICE_PATH"), "soffice", "libreoffice"]

    for c in candidates:
        if not c:
            continue
        if os.path.isabs(c) and os.path.isfile(c):
            return c
        found = shutil.which(c)
        if found:
            return found
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


def check_libreoffice_installed() -> tuple[bool, str]:
    """
    Verify LibreOffice is available for headless conversion.
    Returns (success, message).
    """
    path = _find_libreoffice()
    if path is None:
        return False, (
            "LibreOffice was not found. SlideToObsidian requires LibreOffice for PDF export.\n"
            "  - Install LibreOffice: https://www.libreoffice.org/download\n"
            "  - Or set LIBREOFFICE_PATH to the path of soffice (e.g. soffice.exe on Windows)."
        )
    return True, path


def convert_pptx_to_pdf(pptx_path: str | Path, output_dir: str | Path) -> Path | None:
    """
    Convert a single PPTX file to PDF using LibreOffice headless.
    PDF is written to output_dir with the same stem as the PPTX.
    Returns path to the created PDF, or None on failure.
    """
    pptx_path = Path(pptx_path).resolve()
    output_dir = Path(output_dir).resolve()
    if not pptx_path.is_file():
        return None

    soffice = _find_libreoffice()
    if not soffice:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    # LibreOffice writes to the same directory as the input by default; use a temp dir
    # and then move the PDF to preserve input dir structure
    with tempfile.TemporaryDirectory(prefix="slide2obs_") as tmp:
        tmp_path = Path(tmp)
        try:
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(tmp_path),
                    str(pptx_path),
                ],
                capture_output=True,
                timeout=120,
                check=True,
            )
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
            return None

        pdf_name = pptx_path.stem + ".pdf"
        src_pdf = tmp_path / pdf_name
        if not src_pdf.is_file():
            return None
        dest_pdf = output_dir / pdf_name
        shutil.copy2(src_pdf, dest_pdf)
        return dest_pdf
