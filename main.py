#!/usr/bin/env python3
"""
SlideToObsidian: batch-convert PPTX files into AI-optimized Markdown + PDF pairs for Obsidian.
Uses a parallel worker pool (all CPU cores minus one), markitdown for semantic extraction,
python-pptx for speaker notes, and LibreOffice headless for PDF export.
"""

import argparse
import os
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable

from tqdm import tqdm

# Ensure package is importable when run as script
sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.libreoffice import check_libreoffice_installed
from src.worker import _init_worker, convert_one_worker


# All supported PowerPoint / presentation extensions (LibreOffice can convert these to PDF)
POWERPOINT_EXTENSIONS = {".pptx", ".ppt", ".pot", ".potx", ".pps", ".ppsx"}


def find_powerpoint_files(root: Path) -> list[Path]:
    """Recurse under root and collect all PowerPoint file paths (case-insensitive)."""
    out: list[Path] = []
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in POWERPOINT_EXTENSIONS:
            out.append(p.resolve())
    return sorted(out)


def find_pptx_files(root: Path) -> list[Path]:
    """Alias for find_powerpoint_files (finds .pptx, .ppt, .pot, .potx, .pps, .ppsx)."""
    return find_powerpoint_files(root)


def run_conversion(
    input_path: Path,
    output_base: Path,
    *,
    mirror_structure: bool = True,
    progress_callback: Callable[[int, int, Any], None] | None = None,
) -> tuple[int, list[tuple[str, str]]]:
    """
    Run batch conversion. Returns (success_count, failed_list).
    failed_list is list of (path, error_string).
    If progress_callback is given, call it as (current, total, result) for each completed file.
    """
    pptx_files = find_powerpoint_files(input_path) if input_path.is_dir() else [input_path]
    if not pptx_files:
        return 0, []

    ok, _ = check_libreoffice_installed()
    if not ok:
        raise RuntimeError("LibreOffice was not found. Please install it or set LIBREOFFICE_PATH.")

    output_base = output_base.resolve()
    output_base.mkdir(parents=True, exist_ok=True)
    max_workers = max(1, (os.cpu_count() or 2) - 1)
    max_workers = min(max_workers, len(pptx_files))

    if mirror_structure and input_path.is_dir():
        def out_dir_for(p: Path) -> Path:
            try:
                rel = p.parent.relative_to(input_path)
                return output_base / rel
            except ValueError:
                return output_base
        tasks = [(str(p), str(out_dir_for(p))) for p in pptx_files]
    else:
        tasks = [(str(p), str(output_base)) for p in pptx_files]

    success_count = 0
    failed: list[tuple[str, str]] = []
    total = len(tasks)
    completed = 0

    with ProcessPoolExecutor(max_workers=max_workers, initializer=_init_worker) as executor:
        futures = {executor.submit(convert_one_worker, path, out): (path, out) for path, out in tasks}
        for fut in as_completed(futures):
            path = futures[fut][0]
            result = None
            try:
                result = fut.result()
                if result.get("success"):
                    success_count += 1
                else:
                    failed.append((result.get("path", path), result.get("error", "Unknown error")))
            except Exception as e:
                failed.append((path, str(e)))
            completed += 1
            if progress_callback:
                progress_callback(completed, total, result)
    return success_count, failed


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert PPTX files to AI-optimized Markdown + PDF pairs for Obsidian.",
    )
    parser.add_argument(
        "input",
        type=Path,
        nargs="?",
        default=None,
        help="Input file or directory (optional; use --gui to run GUI instead)",
    )
    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=None,
        help="Output directory (default: 'out/' or mirror under input dir)",
    )
    parser.add_argument(
        "--no-mirror",
        action="store_true",
        help="Put all outputs in one directory (default: mirror input structure when input is a dir)",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the GUI instead of CLI",
    )
    args = parser.parse_args()

    if args.gui:
        from gui import run_app
        return run_app()

    if args.input is None:
        parser.error("input path is required (or use --gui)")
    input_path = args.input.resolve()
    if not input_path.exists():
        print("Error: input path does not exist.", file=sys.stderr)
        return 1
    if input_path.is_file() and input_path.suffix.lower() not in POWERPOINT_EXTENSIONS:
        print("Error: input file is not a supported PowerPoint format (.pptx, .ppt, .pot, .potx, .pps, .ppsx).", file=sys.stderr)
        return 1

    output_base = args.output.resolve() if args.output else (Path.cwd() / "out")
    mirror_structure = not args.no_mirror if input_path.is_dir() else False

    try:
        pbar: tqdm | None = None
        def on_progress(current: int, total: int, result: None) -> None:
            nonlocal pbar
            if pbar is None:
                pbar = tqdm(total=total, unit="file", desc="Converting")
            pbar.update(1)
        success_count, failed = run_conversion(
            input_path, output_base,
            mirror_structure=mirror_structure,
            progress_callback=on_progress,
        )
        if pbar:
            pbar.close()
    except RuntimeError as e:
        print(str(e), file=sys.stderr)
        return 1

    print(f"\nDone: {success_count} converted, {len(failed)} failed.")
    if failed:
        for path, err in failed:
            print(f"  FAILED: {path}\n    {err}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
