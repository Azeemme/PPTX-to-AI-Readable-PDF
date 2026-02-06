"""
Single-file conversion pipeline: PowerPoint → PDF (LibreOffice) + MD (markitdown + builder).
Supports .pptx, .ppt, .pot, .potx, .pps, .ppsx. Uses a shared MarkItDown instance when provided.
"""

from pathlib import Path
from typing import Any

from markitdown import MarkItDown

from .libreoffice import convert_pptx_to_pdf
from .markdown_builder import build_markdown, split_markdown_by_slides
from .pdf_metadata import get_pdf_page_count, set_pdf_metadata
from .pptx_utils import get_slide_count, get_speaker_notes


def convert_one(
    pptx_path: str | Path,
    output_dir: str | Path,
    markitdown_instance: MarkItDown | None = None,
) -> dict[str, Any]:
    """
    Convert one PowerPoint file to PDF + MD in output_dir. Output files use the same stem.
    Supports .pptx, .ppt, .pot, .potx, .pps, .ppsx. Returns {"success": bool, "path": str, "error": str | None}.
    """
    pptx_path = Path(pptx_path).resolve()
    output_dir = Path(output_dir).resolve()
    stem = pptx_path.stem

    if markitdown_instance is None:
        markitdown_instance = MarkItDown()

    try:
        # 1) PDF via LibreOffice (works for all supported formats)
        pdf_path = convert_pptx_to_pdf(pptx_path, output_dir)
        if pdf_path is None:
            return {"success": False, "path": str(pptx_path), "error": "PDF conversion failed (LibreOffice)"}
        set_pdf_metadata(pdf_path, title=stem, source=str(pptx_path))

        # 2) Slide count: python-pptx (only .pptx) or fallback to PDF page count
        slide_count = get_slide_count(pptx_path)
        if slide_count <= 0:
            slide_count = get_pdf_page_count(pdf_path)
        if slide_count <= 0:
            return {"success": False, "path": str(pptx_path), "error": "No slides found or invalid file"}

        # 3) Speaker notes: python-pptx only supports .pptx
        speaker_notes = get_speaker_notes(pptx_path)
        if len(speaker_notes) < slide_count:
            speaker_notes = speaker_notes + [""] * (slide_count - len(speaker_notes))
        speaker_notes = speaker_notes[:slide_count]

        # 4) Semantic body via markitdown (may not support all formats; fallback to empty)
        try:
            result = markitdown_instance.convert(str(pptx_path))
            full_md_body = (result.text_content or "").strip()
        except Exception:
            full_md_body = ""

        # 5) Split markitdown body by slides
        body_sections = split_markdown_by_slides(full_md_body, slide_count)

        # 6) Build Obsidian MD
        md_content = build_markdown(
            pdf_basename=stem,
            slide_count=slide_count,
            speaker_notes=speaker_notes,
            body_sections=body_sections,
            title=stem,
        )
        md_path = output_dir / f"{stem}.md"
        md_path.write_text(md_content, encoding="utf-8")

        return {"success": True, "path": str(pptx_path), "error": None}
    except Exception as e:
        return {"success": False, "path": str(pptx_path), "error": str(e)}
