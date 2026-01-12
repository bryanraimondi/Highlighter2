import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from docx import Document

from parser import extract_text_from_docx, extract_metadata, extract_ecs_rows
from excel_io import read_master, append_and_dedup, to_excel_bytes, COLUMNS


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LB6 Shift Reports → Master Excel")
        self.geometry("860x520")

        self.docx_paths = []
        self.master_path = ""
        self.output_dir = ""
        self.assumed_year = tk.IntVar(value=datetime.utcnow().year)

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        # ── Header row (title on left, author on right) ──────────────
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", **pad)

        header = ttk.Label(
            header_frame,
            text="LB6 Shift Reports (.docx) → Master Excel (append + dedup)",
            font=("Segoe UI", 14)
        )
        header.pack(side="left", anchor="w")

        author = ttk.Label(
            header_frame,
            text="Author: Bryan Raimondi",
            font=("Segoe UI", 9),
            foreground="#555555"
        )
        author.pack(side="right", anchor="e")

        # ── Main frame ───────────────────────────────────────────────
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        # DOCX selection
        row1 = ttk.Frame(frm)
        row1.pack(fill="x", **pad)
        ttk.Button(row1, text="Select DOCX file(s)", command=self.pick_docx).pack(side="left")
        self.docx_lbl = ttk.Label(row1, text="No files selected")
        self.docx_lbl.pack(side="left", padx=10)

        # Master selection
        row2 = ttk.Frame(frm)
        row2.pack(fill="x", **pad)
        ttk.Button(row2, text="Select Master.xlsx (optional)", command=self.pick_master).pack(side="left")
        self.master_lbl = ttk.Label(row2, text="(optional)")
        self.master_lbl.pack(side="left", padx=10)

        # Output dir
        row3 = ttk.Frame(frm)
        row3.pack(fill="x", **pad)
        ttk.Button(row3, text="Select output folder", command=self.pick_output).pack(side="left")
        self.out_lbl = ttk.Label(row3, text="(required)")
        self.out_lbl.pack(side="left", padx=10)

        # Assumed year
        row4 = ttk.Frame(frm)
        row4.pack(fill="x", **pad)
        ttk.Label(row4, text="Assumed year (if Word file has no year):").pack(side="left")
        ttk.Spinbox(
            row4,
            from_=2020,
            to=2035,
            textvariable=self.assumed_year,
            width=8
        ).pack(side="left", padx=10)

        # Process button
        row5 = ttk.Frame(frm)
        row5.pack(fill="x", **pad)
        ttk.Button(row5, text="Process", command=self.process, style="Accent.TButton").pack(side="left")

        # Log box
        ttk.Label(frm, text="Log:").pack(anchor="w", **pad)
        self.log = tk.Text(frm, height=16, wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Button style
        try:
            s = ttk.Style()
            s.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        except Exception:
            pass

    def write_log(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def pick_docx(self):
        paths = filedialog.askopenfilenames(
            title="Select Shift Report DOCX file(s)",
            filetypes=[("Word files", "*.docx")],
        )
        if paths:
            self.docx_paths = list(paths)
            self.docx_lbl.config(text=f"{len(self.docx_paths)} file(s) selected")

    def pick_master(self):
        path = filedialog.askopenfilename(
            title="Select Master.xlsx (optional)",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.master_path = path
            self.master_lbl.config(text=os.path.basename(path))

    def pick_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir = path
            self.out_lbl.config(text=path)

    def process(self):
        if not self.docx_paths:
            messagebox.showerror("Error", "Please select at least one DOCX file.")
            return
        if not self.output_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        self.log.delete("1.0", "end")
        self.write_log("Starting processing...")

        master_bytes = None
        if self.master_path and os.path.exists(self.master_path):
            with open(self.master_path, "rb") as f:
                master_bytes = f.read()

        master_df = read_master(master_bytes)
        before = len(master_df)

        all_new = []
        errors = 0

        for p in self.docx_paths:
            name = os.path.basename(p)
            try:
                self.write_log(f"Processing: {name}")
                doc = Document(p)
                text = extract_text_from_docx(doc)
                meta = extract_metadata(text, assumed_year=int(self.assumed_year.get()))
                rows = extract_ecs_rows(text)

                if not rows:
                    self.write_log(f"  ⚠️ No ECS lines found in {name}")
                    continue

                now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
                out_rows = []
                for ecs_base, item in rows:
                    ecs_full = f"{ecs_base}{item}"
                    out_rows.append({
                        "ECS_CODE_FULL": ecs_full,
                        "ECS_BASE": ecs_base,
                        "ITEM": item,
                        "WORK_DATE": meta.work_date,
                        "SUPERVISOR": meta.supervisor,
                        "SUPERINTENDENT": meta.superintendent,
                        "SOURCE_FILE": name,
                        "INGESTED_AT": now,
                    })

                df_new = pd.DataFrame(out_rows, columns=COLUMNS)
                all_new.append(df_new)

            except Exception as e:
                errors += 1
                self.write_log(f"  ❌ ERROR in {name}: {e}")

        if not all_new:
            messagebox.showwarning(
                "Warning",
                "No ECS rows could be extracted from the selected files."
            )
            return

        new_df = pd.concat(all_new, ignore_index=True)
        updated = append_and_dedup(master_df, new_df)
        after = len(updated)
        delta = after - before

        out_path = os.path.join(self.output_dir, "Master_updated.xlsx")
        with open(out_path, "wb") as f:
            f.write(to_excel_bytes(updated, sheet_name="MASTER"))

        self.write_log("")
        self.write_log(
            f"Completed successfully. Rows before: {before} | "
            f"after: {after} | net added: {delta}"
        )
        if errors:
            self.write_log(f"Files with errors: {errors}")

        messagebox.showinfo(
            "Done",
            f"Master file updated successfully.\n\nSaved at:\n{out_path}"
        )


if __name__ == "__main__":
    App().mainloop()
