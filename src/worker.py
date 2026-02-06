"""
Worker for ProcessPoolExecutor: one MarkItDown per process, reused for all tasks.
Each task converts one PPTX; exceptions are caught so one bad file doesn't kill the batch.
"""

from pathlib import Path
from typing import Any

from markitdown import MarkItDown

# Module-level MarkItDown created in worker init; reused for all tasks in this process
_worker_markitdown: MarkItDown | None = None


def _init_worker() -> None:
    """Called once per worker process; create a single MarkItDown instance."""
    global _worker_markitdown
    _worker_markitdown = MarkItDown()


def convert_one_worker(
    pptx_path: str,
    output_dir: str,
) -> dict[str, Any]:
    """
    Convert a single PPTX to PDF + MD. Uses the process-local MarkItDown.
    Always returns a result dict; never raises (one corrupt file doesn't crash the batch).
    """
    try:
        from .converter import convert_one
        md = _worker_markitdown if _worker_markitdown is not None else None
        return convert_one(
            Path(pptx_path),
            Path(output_dir),
            markitdown_instance=md,
        )
    except Exception as e:
        return {
            "success": False,
            "path": pptx_path,
            "error": str(e),
        }
