# -*- coding: utf-8 -*-
"""Desktop UI — ILE vs Topos, không cần Excel / Office."""

from __future__ import annotations

import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from engine import OUTPUT_SUFFIX, CompareError, compare_workbook

BG = "#f4f6f8"
NAVY = "#1f5179"
DROP_BG = "#ffffff"
DROP_HINT = "#6b7280"
EXCEL_TYPES = (("Excel (*.xlsx)", "*.xlsx"), ("All files (*.*)", "*.*"))


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _xlsx_paths(raw_paths) -> list[str]:
    out = []
    for raw in raw_paths or []:
        path = Path(str(raw).strip().strip('"'))
        if path.suffix.lower() == ".xlsx" and path.exists():
            out.append(str(path.resolve()))
    return out


class FileDropPanel(tk.Frame):
    def __init__(self, master, title, hint, multiple=True, browse_title="Chọn file"):
        super().__init__(master, bg=BG, highlightthickness=1, highlightbackground="#cbd5e1")
        self.multiple = multiple
        self.browse_title = browse_title
        self.files: list[str] = []
        tk.Label(self, text=title, font=("Segoe UI Semibold", 10), bg=BG, anchor="w").pack(
            fill="x", padx=8, pady=(8, 0)
        )
        tk.Label(self, text=hint, font=("Segoe UI", 8), fg=DROP_HINT, bg=BG, anchor="w").pack(
            fill="x", padx=8, pady=(0, 4)
        )
        body = tk.Frame(self, bg=DROP_BG)
        body.pack(fill="both", expand=True, padx=8, pady=4)
        list_frame = tk.Frame(body, bg=DROP_BG)
        list_frame.pack(side="left", fill="both", expand=True)
        self.listbox = tk.Listbox(
            list_frame,
            selectmode=(tk.EXTENDED if multiple else tk.BROWSE),
            font=("Consolas", 9),
            bg=DROP_BG,
            activestyle="dotbox",
            height=(5 if multiple else 3),
        )
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        btns = tk.Frame(body, bg=DROP_BG)
        btns.pack(side="right", fill="y", padx=6)
        ttk.Button(btns, text="Thêm file…", command=self.browse, width=12).pack(pady=2)
        ttk.Button(btns, text="Xóa chọn", command=self.remove_selected, width=12).pack(pady=2)
        ttk.Button(btns, text="Xóa hết", command=self.clear, width=12).pack(pady=2)

    def browse(self):
        if self.multiple:
            paths = filedialog.askopenfilenames(title=self.browse_title, filetypes=EXCEL_TYPES)
        else:
            path = filedialog.askopenfilename(title=self.browse_title, filetypes=EXCEL_TYPES)
            paths = (path,) if path else ()
        if paths:
            self.add_paths(paths)

    def add_paths(self, paths) -> int:
        added = 0
        for path in _xlsx_paths(paths):
            if not self.multiple:
                self.files = [path]
                added = 1
                break
            if path not in self.files:
                self.files.append(path)
                added += 1
        self._refresh()
        return added

    def remove_selected(self):
        for idx in reversed(list(self.listbox.curselection())):
            if 0 <= idx < len(self.files):
                del self.files[idx]
        self._refresh()

    def clear(self):
        self.files.clear()
        self._refresh()

    def selected_or_first(self):
        sel = self.listbox.curselection()
        if sel:
            return self.files[sel[0]]
        return self.files[0] if self.files else None

    def _refresh(self):
        self.listbox.delete(0, tk.END)
        for path in self.files:
            self.listbox.insert(tk.END, path)


class CompareApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ILE vs Topos Compare 1.1 — Portable (không cần Excel)")
        self.geometry("920x720")
        self.minsize(780, 580)
        self.configure(bg=BG)
        self.last_output = None
        self.busy = False
        self._build()

    def _build(self):
        header = tk.Frame(self, bg=NAVY)
        header.pack(fill="x")
        tk.Label(
            header,
            text="ILE vs Topos Compare 1.1",
            font=("Segoe UI Semibold", 16),
            fg="white",
            bg=NAVY,
        ).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(
            header,
            text="Key: Ngày · Location/Tên Cửa Hàng · Type · Item No/Mã Nội Bộ   ·   Lệch = Topos − Item Ledger",
            font=("Segoe UI", 10),
            fg="#d6e5fc",
            bg=NAVY,
        ).pack(anchor="w", padx=16, pady=(2, 12))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=4)
        tab_one, tab_two = tk.Frame(self.notebook, bg=BG), tk.Frame(self.notebook, bg=BG)
        self.notebook.add(tab_one, text="  Một file (2 sheet)  ")
        self.notebook.add(tab_two, text="  Hai file riêng  ")
        self._build_one_tab(tab_one)
        self._build_two_tab(tab_two)

        self.status = tk.Label(self, text="Sẵn sàng.", anchor="w", bg=BG, font=("Segoe UI", 9))
        self.status.pack(fill="x", padx=8)
        log_frame = tk.Frame(self, bg=BG)
        log_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.log = tk.Text(
            log_frame, height=10, font=("Consolas", 9), bg="#111827", fg="#e5e7eb", wrap="word", state="disabled"
        )
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self._write_log("Sẵn sàng. Không cần Microsoft Excel.")
        self._write_log("Chỉ liệt kê key lệch. File kết quả: '<tên gốc> - ILE vs Topos.xlsx'.")

    def _build_one_tab(self, parent):
        tk.Label(
            parent,
            text=(
                "Một file .xlsx có sheet Item Ledger Entries và Topos.\n"
                "Quantity được cộng dồn theo từng key rồi so sánh."
            ),
            font=("Segoe UI", 9),
            justify="left",
            bg=BG,
        ).pack(anchor="w", padx=12, pady=(10, 6))
        self.one_files = FileDropPanel(
            parent,
            "File dữ liệu gốc",
            "Chọn một hoặc nhiều file (mỗi file đủ 2 sheet).",
            multiple=True,
            browse_title="Chọn file Excel (ILE + Topos)",
        )
        self.one_files.pack(fill="both", expand=True, padx=12, pady=4)
        actions = tk.Frame(parent, bg=BG)
        actions.pack(fill="x", padx=12, pady=(4, 10))
        self.one_run_btn = ttk.Button(actions, text="Chạy đối chiếu", command=lambda: self.run_compare("one"))
        self.one_run_btn.pack(side="left")
        ttk.Button(actions, text="Mở kết quả", command=self.open_result).pack(side="left", padx=6)
        ttk.Button(actions, text="Mở thư mục", command=self.open_folder).pack(side="left")

    def _build_two_tab(self, parent):
        tk.Label(
            parent,
            text="File Item Ledger và file Topos tách riêng. Có thể chọn nhiều cặp bằng cách chạy lần lượt.",
            font=("Segoe UI", 9),
            justify="left",
            bg=BG,
        ).pack(anchor="w", padx=12, pady=(10, 6))
        self.ile_files = FileDropPanel(
            parent,
            "1. File Item Ledger Entries",
            "Chọn file chứa sheet Item Ledger (hoặc chỉ có 1 sheet).",
            multiple=False,
            browse_title="Chọn file Item Ledger",
        )
        self.ile_files.pack(fill="x", padx=12, pady=4)
        self.topos_files = FileDropPanel(
            parent,
            "2. File Topos",
            "Chọn file chứa sheet Topos (hoặc chỉ có 1 sheet).",
            multiple=False,
            browse_title="Chọn file Topos",
        )
        self.topos_files.pack(fill="x", padx=12, pady=4)
        actions = tk.Frame(parent, bg=BG)
        actions.pack(fill="x", padx=12, pady=(4, 10))
        self.two_run_btn = ttk.Button(actions, text="Chạy đối chiếu", command=lambda: self.run_compare("two"))
        self.two_run_btn.pack(side="left")
        ttk.Button(actions, text="Mở kết quả", command=self.open_result).pack(side="left", padx=6)
        ttk.Button(actions, text="Mở thư mục", command=self.open_folder).pack(side="left")

    def _write_log(self, message, kind=""):
        stamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"ERR": "    ERR: ", "OK": "    OK "}.get(kind, "")
        line = f"[{stamp}] {prefix}{message}\n"

        def append():
            self.log.configure(state="normal")
            self.log.insert(tk.END, line)
            self.log.see(tk.END)
            self.log.configure(state="disabled")

        if threading.current_thread() is threading.main_thread():
            append()
        else:
            self.after(0, append)

    def _current_data_panel(self):
        return self.ile_files if self.notebook.index(self.notebook.select()) == 1 else self.one_files

    def open_result(self):
        target = self.last_output
        if not target:
            src_path = self._current_data_panel().selected_or_first()
            if src_path:
                candidate = Path(src_path).with_name(Path(src_path).stem + OUTPUT_SUFFIX)
                if candidate.exists():
                    target = str(candidate)
        if not target or not Path(target).exists():
            messagebox.showinfo("Chưa có file", "Chưa có file kết quả. Hãy chạy đối chiếu trước.")
            return
        os.startfile(target)

    def open_folder(self):
        src_path = self._current_data_panel().selected_or_first() or self.last_output
        if not src_path:
            messagebox.showinfo("Chưa có file", "Chưa có file.")
            return
        if Path(src_path).exists():
            os.system(f'explorer /select,"{src_path}"')
        else:
            os.startfile(str(Path(src_path).parent))

    def run_compare(self, mode="one"):
        if self.busy:
            return
        if mode == "two":
            ile = self.ile_files.selected_or_first()
            topos = self.topos_files.selected_or_first()
            if not ile:
                messagebox.showwarning("Chưa có file", "Hãy chọn file Item Ledger.")
                return
            if not topos:
                messagebox.showwarning("Chưa có file", "Hãy chọn file Topos.")
                return
            jobs = [(ile, topos)]
            btn = self.two_run_btn
        else:
            files = list(self.one_files.files)
            if not files:
                messagebox.showwarning("Chưa có file", "Hãy thêm ít nhất 1 file .xlsx.")
                return
            jobs = [(path, None) for path in files]
            btn = self.one_run_btn
        self.busy = True
        btn.configure(state="disabled")
        self.status.configure(text="Đang chạy…")
        threading.Thread(target=self._run_worker, args=(jobs, btn), daemon=True).start()

    def _run_worker(self, jobs, btn):
        last_error, ok = None, 0
        for ile_path, topos_path in jobs:
            try:
                self._write_log(f"Đang xử lý: {ile_path}")
                if topos_path:
                    self._write_log(f"Topos: {topos_path}")
                result = compare_workbook(ile_path, log=self._write_log, topos_path=topos_path)
                self.last_output = result["OutputPath"]
                self._write_log(
                    f"Xong {Path(ile_path).name}: {result['MismatchRows']} key lệch "
                    f"(cả hai {result['QtyDiffBoth']}, chỉ ILE {result['OnlyIle']}, "
                    f"chỉ Topos {result['OnlyTopos']})",
                    "OK",
                )
                ok += 1
            except Exception as exc:
                last_error = exc
                self._write_log(str(exc), "ERR")

        def finish():
            self.busy = False
            btn.configure(state="normal")
            if last_error and ok == 0:
                self.status.configure(text="Lỗi. Xem nhật ký.")
                messagebox.showerror("Lỗi", str(last_error))
            else:
                self.status.configure(text="Đã tạo file ILE vs Topos.")
                messagebox.showinfo("Xong", f"Đã xử lý {ok}/{len(jobs)} file.\nKhông cần Excel.")

        self.after(0, finish)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    xlsx_args = [a for a in args if a.lower().endswith(".xlsx")]
    app = CompareApp()
    if xlsx_args:
        app.one_files.add_paths(xlsx_args)
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        text = "App khong khoi dong duoc:\n" + "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("ILE vs Topos Compare", text[:2000])
            root.destroy()
        except Exception:
            pass
        raise
