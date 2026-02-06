#!/usr/bin/env python3
"""
SlideToObsidian GUI: pick a directory, process all PPTX files with one button.
"""

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# Ensure package is importable
sys.path.insert(0, str(Path(__file__).resolve().parent))

from main import find_powerpoint_files, run_conversion
from src.libreoffice import check_libreoffice_installed


def run_app() -> int:
    root = tk.Tk()
    app = SlideToObsidianApp(root)
    app.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    root.title("SlideToObsidian")
    root.minsize(420, 380)
    root.geometry("560x420")
    root.mainloop()
    return 0


class SlideToObsidianApp(tk.Frame):
    def __init__(self, master: tk.Tk, **kwargs: object) -> None:
        super().__init__(master, **kwargs)
        self.input_dir: Path | None = None
        self.progress_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        # Directory row
        dir_frame = ttk.Frame(self)
        dir_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(dir_frame, text="Directory:").pack(side=tk.LEFT, padx=(0, 6))
        self.dir_var = tk.StringVar()
        self.dir_entry = ttk.Entry(dir_frame, textvariable=self.dir_var, state="readonly")
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.browse_btn = ttk.Button(dir_frame, text="Browse…", command=self._on_browse)
        self.browse_btn.pack(side=tk.RIGHT)

        # File list area
        list_label = ttk.Label(self, text="PowerPoint files in this directory (and subfolders):")
        list_label.pack(anchor=tk.W, pady=(8, 2))
        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.file_list = tk.Text(list_frame, height=8, wrap=tk.WORD, state="disabled", font=("Segoe UI", 9))
        scroll = ttk.Scrollbar(list_frame, command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=scroll.set)
        self.file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Process button
        self.process_btn = ttk.Button(self, text="Process", command=self._on_process, state="disabled")
        self.process_btn.pack(fill=tk.X, pady=(0, 8))

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(self, variable=self.progress_var, maximum=100)
        self.progress.pack(fill=tk.X, pady=(0, 8))

        # Log area
        log_label = ttk.Label(self, text="Log:")
        log_label.pack(anchor=tk.W, pady=(0, 2))
        log_frame = ttk.Frame(self)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, height=6, wrap=tk.WORD, state="disabled", font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _log(self, msg: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _on_browse(self) -> None:
        path = filedialog.askdirectory(title="Select folder containing PPTX files")
        if path:
            self.dir_var.set(path)
            self.input_dir = Path(path)
            file_list = find_powerpoint_files(self.input_dir)
            self.file_list.configure(state="normal")
            self.file_list.delete("1.0", tk.END)
            if file_list:
                self.file_list.insert(tk.END, f"{len(file_list)} file(s) found:\n")
                for p in file_list:
                    self.file_list.insert(tk.END, f"  • {p.name}\n")
                self.process_btn.configure(state="normal")
            else:
                self.file_list.insert(tk.END, "No PowerPoint files (.pptx, .ppt, .pot, .potx, .pps, .ppsx) found.")
            self.file_list.configure(state="disabled")
            self._log(f"Selected: {path} — {len(file_list)} PowerPoint file(s)")

    def _on_process(self) -> None:
        if not self.input_dir or not self.input_dir.is_dir():
            messagebox.showerror("Error", "Please select a directory first.")
            return
        ok, msg = check_libreoffice_installed()
        if not ok:
            messagebox.showerror("LibreOffice required", msg)
            return
        file_list = find_powerpoint_files(self.input_dir)
        if not file_list:
            messagebox.showwarning("No files", "No PowerPoint files found in the selected directory.")
            return
        self.process_btn.configure(state="disabled")
        self.browse_btn.configure(state="disabled")
        self.progress_var.set(0)
        self.progress.configure(maximum=100, mode="determinate")
        self._log("Starting conversion…")
        self.worker_thread = threading.Thread(target=self._run_conversion, daemon=True)
        self.worker_thread.start()
        self._drain_queue()

    def _run_conversion(self) -> None:
        try:
            output_base = self.input_dir / "out"
            total = [0]
            def on_progress(current: int, t: int, result: object) -> None:
                total[0] = t
                self.progress_queue.put(("progress", current, t))
            success_count, failed = run_conversion(
                self.input_dir,
                output_base,
                mirror_structure=True,
                progress_callback=on_progress,
            )
            self.progress_queue.put(("done", success_count, failed, str(output_base)))
        except Exception as e:
            self.progress_queue.put(("error", str(e)))

    def _drain_queue(self) -> None:
        try:
            while True:
                msg = self.progress_queue.get_nowait()
                if msg[0] == "progress":
                    _, current, total = msg
                    if total:
                        self.progress_var.set(100.0 * current / total)
                elif msg[0] == "done":
                    _, success_count, failed, output_base = msg
                    self.progress_var.set(100)
                    self._log(f"Done: {success_count} converted, {len(failed)} failed.")
                    self._log(f"Output folder: {output_base}")
                    for path, err in failed:
                        self._log(f"  FAILED: {path} — {err}")
                    self.process_btn.configure(state="normal")
                    self.browse_btn.configure(state="normal")
                    messagebox.showinfo("Complete", f"Converted {success_count} file(s). Output: {output_base}")
                    return
                elif msg[0] == "error":
                    self._log(f"Error: {msg[1]}")
                    self.process_btn.configure(state="normal")
                    self.browse_btn.configure(state="normal")
                    messagebox.showerror("Error", msg[1])
                    return
        except queue.Empty:
            pass
        self.after(150, self._drain_queue)
