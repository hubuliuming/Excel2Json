# -*- coding: utf-8 -*-
"""
Excel Sheet Exporter Launcher (no .bat required)

Usage:
1. Put this file in the same folder as excel_sheet_exporter.py
2. Double-click this .pyw file, or run with: python excel_export_launcher.pyw
3. Click "Run Export"

Notes:
- Designed for legacy layout:
    row 1 = field names
    row 2 = types
    row 3+ = data
- A sheet is exportable only if row 1 contains a column named "ClassName"
- If client output folder does not exist, JSON copy is skipped and export still succeeds
"""

import os
import sys
import threading
import traceback
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Excel Export Launcher"

def find_python_command() -> list[str]:
    candidates = [
        ["py", "-3"],
        [sys.executable],
        ["python"],
    ]
    for cmd in candidates:
        try:
            result = subprocess.run(
                cmd + ["-c", "print('ok')"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                shell=False,
            )
            if result.returncode == 0 and "ok" in result.stdout:
                return cmd
        except Exception:
            pass
    raise RuntimeError("Python interpreter not found. Please install Python and add it to PATH.")

def check_openpyxl(py_cmd: list[str]) -> str:
    result = subprocess.run(
        py_cmd + ["-c", "import openpyxl; print(openpyxl.__version__)"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        shell=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            "openpyxl is missing.\n\n"
            f"Install with:\n{' '.join(py_cmd)} -m pip install openpyxl\n\n"
            f"Details:\n{result.stderr.strip()}"
        )
    return result.stdout.strip()

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("860x680")
        self.root.minsize(760, 560)

        self.cur_dir = Path(__file__).resolve().parent
        self.script_path = self.cur_dir / "excel_sheet_exporter.py"

        self.excel_folder_var = tk.StringVar(value=str((self.cur_dir / ".." / "Excel").resolve()))
        self.json_folder_var = tk.StringVar(value=str((self.cur_dir / "json").resolve()))
        self.cs_folder_var = tk.StringVar(value=str((self.cur_dir / "cs").resolve()))
        self.client_json_folder_var = tk.StringVar(value=str((self.cur_dir / ".." / ".." / "client" / "JsonData").resolve()))
        self.namespace_var = tk.StringVar(value="Game.Config")
        self.copy_to_client_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self):
        root = self.root
        pad = {"padx": 10, "pady": 6}

        main = ttk.Frame(root)
        main.pack(fill="both", expand=True)

        top = ttk.Frame(main)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(top, text="Excel Export Launcher", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(
            top,
            text="Legacy layout: row1 = headers, row2 = types, row3+ = data. No .bat required.",
        ).pack(anchor="w", pady=(4, 0))

        form = ttk.Frame(main)
        form.pack(fill="x", padx=10)

        self._path_row(form, 0, "Exporter Script", str(self.script_path), None, readonly=True)
        self._path_row(form, 1, "Excel Folder", self.excel_folder_var, self.pick_dir)
        self._path_row(form, 2, "JSON Folder", self.json_folder_var, self.pick_dir)
        self._path_row(form, 3, "C# Folder", self.cs_folder_var, self.pick_dir)
        self._path_row(form, 4, "Client JSON Folder", self.client_json_folder_var, self.pick_dir)

        ttk.Label(form, text="Namespace").grid(row=5, column=0, sticky="w", **pad)
        ttk.Entry(form, textvariable=self.namespace_var).grid(row=5, column=1, sticky="ew", **pad)
        form.grid_columnconfigure(1, weight=1)

        ttk.Checkbutton(
            form,
            text="Copy JSON to Client folder if it exists",
            variable=self.copy_to_client_var,
        ).grid(row=6, column=1, sticky="w", **pad)

        btns = ttk.Frame(main)
        btns.pack(fill="x", padx=10, pady=(8, 4))

        self.run_btn = ttk.Button(btns, text="Run Export", command=self.run_export)
        self.run_btn.pack(side="left")

        ttk.Button(btns, text="Open Tool Folder", command=self.open_tool_folder).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Clear Log", command=self.clear_log).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Save Config", command=self.save_config).pack(side="right")

        ttk.Label(main, textvariable=self.status_var).pack(anchor="w", padx=12, pady=(4, 2))

        log_frame = ttk.Frame(main)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.log = tk.Text(log_frame, wrap="word", font=("Consolas", 10))
        self.log.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scroll.pack(side="right", fill="y")
        self.log.config(yscrollcommand=scroll.set)

        self.load_config()
        self.write_log("[INFO] Launcher ready.\n")

    def _path_row(self, parent, row, label, variable_or_text, browse_cmd, readonly=False):
        pad = {"padx": 10, "pady": 6}
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)

        if isinstance(variable_or_text, tk.StringVar):
            entry = ttk.Entry(parent, textvariable=variable_or_text)
        else:
            temp = tk.StringVar(value=variable_or_text)
            entry = ttk.Entry(parent, textvariable=temp)

        if readonly:
            entry.state(["readonly"])
        entry.grid(row=row, column=1, sticky="ew", **pad)

        if browse_cmd:
            ttk.Button(parent, text="Browse...", command=lambda v=variable_or_text: browse_cmd(v)).grid(
                row=row, column=2, sticky="ew", **pad
            )

    def write_log(self, text: str):
        self.log.insert("end", text)
        self.log.see("end")
        self.root.update_idletasks()

    def clear_log(self):
        self.log.delete("1.0", "end")

    def set_status(self, text: str):
        self.status_var.set(text)
        self.root.update_idletasks()

    def pick_dir(self, var: tk.StringVar):
        initial = var.get().strip() or str(self.cur_dir)
        path = filedialog.askdirectory(initialdir=initial)
        if path:
            var.set(path)

    def open_tool_folder(self):
        folder = str(self.cur_dir)
        try:
            os.startfile(folder)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Cannot open folder:\n{folder}\n\n{e}")

    def config_path(self) -> Path:
        return self.cur_dir / "excel_export_launcher_config.txt"

    def save_config(self):
        data = {
            "excel_folder": self.excel_folder_var.get().strip(),
            "json_folder": self.json_folder_var.get().strip(),
            "cs_folder": self.cs_folder_var.get().strip(),
            "client_json_folder": self.client_json_folder_var.get().strip(),
            "namespace": self.namespace_var.get().strip(),
            "copy_to_client": "1" if self.copy_to_client_var.get() else "0",
        }
        lines = [f"{k}={v}" for k, v in data.items()]
        self.config_path().write_text("\n".join(lines), encoding="utf-8")
        self.write_log("[INFO] Config saved.\n")

    def load_config(self):
        path = self.config_path()
        if not path.exists():
            return
        try:
            data = {}
            for line in path.read_text(encoding="utf-8").splitlines():
                if "=" in line:
                    k, v = line.split("=", 1)
                    data[k.strip()] = v.strip()
            self.excel_folder_var.set(data.get("excel_folder", self.excel_folder_var.get()))
            self.json_folder_var.set(data.get("json_folder", self.json_folder_var.get()))
            self.cs_folder_var.set(data.get("cs_folder", self.cs_folder_var.get()))
            self.client_json_folder_var.set(data.get("client_json_folder", self.client_json_folder_var.get()))
            self.namespace_var.set(data.get("namespace", self.namespace_var.get()))
            self.copy_to_client_var.set(data.get("copy_to_client", "1") == "1")
        except Exception as e:
            self.write_log(f"[WARN] Failed to load config: {e}\n")

    def validate_inputs(self):
        if not self.script_path.exists():
            raise RuntimeError(f"Exporter script not found:\n{self.script_path}")
        if not Path(self.excel_folder_var.get().strip()).exists():
            raise RuntimeError(f"Excel folder not found:\n{self.excel_folder_var.get().strip()}")

    def run_export(self):
        self.run_btn.config(state="disabled")
        threading.Thread(target=self._run_export_impl, daemon=True).start()

    def _run_export_impl(self):
        try:
            self.set_status("Checking environment...")
            self.write_log("========================================\n")
            self.write_log("Excel Export Launcher\n")
            self.write_log("========================================\n")
            self.validate_inputs()

            py_cmd = find_python_command()
            py_ver = subprocess.run(
                py_cmd + ["--version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                shell=False,
            )
            self.write_log(f"[INFO] Python command: {' '.join(py_cmd)}\n")
            self.write_log(f"[INFO] Python version: {(py_ver.stdout or py_ver.stderr).strip()}\n")

            openpyxl_ver = check_openpyxl(py_cmd)
            self.write_log(f"[INFO] openpyxl version: {openpyxl_ver}\n")
            self.write_log(f"[INFO] Exporter script: {self.script_path}\n")
            self.write_log(f"[INFO] Excel folder: {self.excel_folder_var.get().strip()}\n")
            self.write_log(f"[INFO] JSON folder: {self.json_folder_var.get().strip()}\n")
            self.write_log(f"[INFO] C# folder: {self.cs_folder_var.get().strip()}\n")
            self.write_log(f"[INFO] Namespace: {self.namespace_var.get().strip()}\n")

            json_dir = Path(self.json_folder_var.get().strip())
            cs_dir = Path(self.cs_folder_var.get().strip())
            json_dir.mkdir(parents=True, exist_ok=True)
            cs_dir.mkdir(parents=True, exist_ok=True)

            cmd = py_cmd + [
                str(self.script_path),
                "--excel-folder", self.excel_folder_var.get().strip(),
                "--json-folder", self.json_folder_var.get().strip(),
                "--cs-folder", self.cs_folder_var.get().strip(),
                "--type-row", "2",
                "--header-row", "1",
                "--data-start-row", "3",
                "--namespace", self.namespace_var.get().strip(),
            ]

            client_dir = self.client_json_folder_var.get().strip()
            if self.copy_to_client_var.get() and client_dir and Path(client_dir).exists():
                cmd += ["--copy-json-to", client_dir]
                self.write_log(f"[INFO] Client JSON copy enabled: {client_dir}\n")
            else:
                self.write_log("[WARN] Client JSON copy skipped.\n")

            self.write_log("\n[INFO] Running command:\n")
            self.write_log(" ".join(f'"{x}"' if " " in x else x for x in cmd) + "\n\n")

            self.set_status("Exporting...")
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                shell=False,
                bufsize=1,
            )

            assert process.stdout is not None
            for line in process.stdout:
                self.write_log(line)

            process.wait()
            self.write_log(f"\n[INFO] Exit code: {process.returncode}\n")

            if process.returncode == 0:
                self.set_status("Export completed")
                self.write_log("[SUCCESS] Export completed.\n")
                self.save_config()
                messagebox.showinfo(APP_TITLE, "Export completed.")
            else:
                self.set_status("Export failed")
                messagebox.showerror(APP_TITLE, "Export failed. See the log window for details.")

        except Exception as e:
            self.set_status("Error")
            self.write_log(f"\n[ERROR] {e}\n")
            self.write_log(traceback.format_exc() + "\n")
            messagebox.showerror(APP_TITLE, str(e))
        finally:
            self.run_btn.config(state="normal")

def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
