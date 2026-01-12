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

        header = ttk.Label(self, text="LB6 Shift Reports (.docx) → Master Excel (append + dedup)", font=("Segoe UI", 14))
        header.pack(anchor="w", **pad)

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        # DOCX selection
        row1 = ttk.Frame(frm)
        row1.pack(fill="x", **pad)
        ttk.Button(row1, text="Selecionar DOCX(s)", command=self.pick_docx).pack(side="left")
        self.docx_lbl = ttk.Label(row1, text="Nenhum arquivo selecionado")
        self.docx_lbl.pack(side="left", padx=10)

        # Master selection
        row2 = ttk.Frame(frm)
        row2.pack(fill="x", **pad)
        ttk.Button(row2, text="Selecionar Master.xlsx (opcional)", command=self.pick_master).pack(side="left")
        self.master_lbl = ttk.Label(row2, text="(opcional)")
        self.master_lbl.pack(side="left", padx=10)

        # Output dir
        row3 = ttk.Frame(frm)
        row3.pack(fill="x", **pad)
        ttk.Button(row3, text="Selecionar pasta de saída", command=self.pick_output).pack(side="left")
        self.out_lbl = ttk.Label(row3, text="(obrigatório)")
        self.out_lbl.pack(side="left", padx=10)

        # Assumed year
        row4 = ttk.Frame(frm)
        row4.pack(fill="x", **pad)
        ttk.Label(row4, text="Ano assumido (se o Word não tiver ano):").pack(side="left")
        ttk.Spinbox(row4, from_=2020, to=2035, textvariable=self.assumed_year, width=8).pack(side="left", padx=10)

        # Process button
        row5 = ttk.Frame(frm)
        row5.pack(fill="x", **pad)
        ttk.Button(row5, text="Processar", command=self.process, style="Accent.TButton").pack(side="left")

        # Log box
        ttk.Label(frm, text="Log:").pack(anchor="w", **pad)
        self.log = tk.Text(frm, height=16, wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Simple style
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
            title="Selecione os Shift Reports (.docx)",
            filetypes=[("Word files", "*.docx")],
        )
        if paths:
            self.docx_paths = list(paths)
            self.docx_lbl.config(text=f"{len(self.docx_paths)} arquivo(s) selecionado(s)")

    def pick_master(self):
        path = filedialog.askopenfilename(
            title="Selecione o Master.xlsx (opcional)",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.master_path = path
            self.master_lbl.config(text=os.path.basename(path))

    def pick_output(self):
        path = filedialog.askdirectory(title="Selecione a pasta de saída")
        if path:
            self.output_dir = path
            self.out_lbl.config(text=path)

    def process(self):
        if not self.docx_paths:
            messagebox.showerror("Erro", "Selecione pelo menos 1 arquivo .docx.")
            return
        if not self.output_dir:
            messagebox.showerror("Erro", "Selecione uma pasta de saída.")
            return

        self.log.delete("1.0", "end")
        self.write_log("Iniciando...")

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
                self.write_log(f"Processando: {name}")
                doc = Document(p)
                text = extract_text_from_docx(doc)
                meta = extract_metadata(text, assumed_year=int(self.assumed_year.get()))
                rows = extract_ecs_rows(text)

                if not rows:
                    self.write_log(f"  ⚠️ Nenhuma linha ECS encontrada em {name}")
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
                self.write_log(f"  ❌ ERRO em {name}: {e}")

        if not all_new:
            messagebox.showwarning("Aviso", "Não consegui extrair nenhuma linha dos arquivos selecionados.")
            return

        new_df = pd.concat(all_new, ignore_index=True)
        updated = append_and_dedup(master_df, new_df)
        after = len(updated)
        delta = after - before

        out_path = os.path.join(self.output_dir, "Master_updated.xlsx")
        with open(out_path, "wb") as f:
            f.write(to_excel_bytes(updated, sheet_name="MASTER"))

        self.write_log("")
        self.write_log(f"Concluído. Linhas antes: {before} | depois: {after} | adicionadas (líquido): {delta}")
        if errors:
            self.write_log(f"Arquivos com erro: {errors}")

        messagebox.showinfo("OK", f"Master atualizado!\n\nSalvo em:\n{out_path}")


if __name__ == "__main__":
    App().mainloop()
