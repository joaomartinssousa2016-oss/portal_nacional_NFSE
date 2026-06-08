import os
import sys
import threading
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import io
import contextlib

import pandas as pd
from NovaBusca import main as worker_main

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

EXCEL_PATH = APP_DIR / "empresas.xlsx"
SCRIPT_PATH = APP_DIR / "NovaBusca.py"
WORKER_SCRIPT_PATH = APP_DIR / "NovaBuscaWorker.py"
OUTPUT_DIR = APP_DIR / "notasfiscais"

COLUMNS = ["Nome empresa", "cnpj", "senha", "data_inicio", "data_fim"]


class NovaBuscaUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Portal Nacional NFSe - NovaBusca")
        self.root.geometry("1200x760")
        self.root.minsize(1020, 680)

        self.process = None
        self.is_running = False
        self.progress_running = False
        self.selected_row_indexes = None

        self._build_styles()
        self._build_layout()
        self._load_excel_if_exists()

    def _build_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground="#1f3f5b")
        style.configure("Sub.TLabel", font=("Segoe UI", 10), foreground="#3b5b73")
        style.configure("Card.TFrame", background="#f4f8fb")
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

    def _build_layout(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        header = ttk.Frame(self.root, padding=(16, 12))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="NovaBusca - Painel Visual", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Cadastre empresas, salve no Excel e rode a busca com log em tempo real.",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        form = ttk.LabelFrame(self.root, text="Cadastro", padding=12)
        form.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        for idx in range(10):
            form.columnconfigure(idx, weight=1)

        self.nome_var = tk.StringVar()
        self.cnpj_var = tk.StringVar()
        self.senha_var = tk.StringVar()
        self.data_inicio_var = tk.StringVar()
        self.data_fim_var = tk.StringVar()

        self._add_field(form, "Nome empresa", self.nome_var, 0)
        self._add_field(form, "CNPJ", self.cnpj_var, 2)
        self._add_field(form, "Senha", self.senha_var, 4, show="*")
        self._add_field(form, "Data inicio", self.data_inicio_var, 6)
        self._add_field(form, "Data fim", self.data_fim_var, 8)

        # Seção de formato de download
        format_frame = ttk.LabelFrame(form, text="Formato de download", padding=8)
        format_frame.grid(row=3, column=0, columnspan=10, sticky="ew", pady=(10, 0))

        self.download_xml_var = tk.BooleanVar(value=True)
        self.download_pdf_var = tk.BooleanVar(value=False)
        self.download_pdf_tipo_var = tk.StringVar(value="todos")

        ttk.Checkbutton(format_frame, text="XML", variable=self.download_xml_var).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(format_frame, text="PDF", variable=self.download_pdf_var).pack(side="left")
        ttk.Label(format_frame, text="Tipo PDF:").pack(side="left", padx=(16, 6))
        ttk.Radiobutton(format_frame, text="Todos", variable=self.download_pdf_tipo_var, value="todos").pack(side="left", padx=(0, 8))
        ttk.Radiobutton(format_frame, text="Emitidas", variable=self.download_pdf_tipo_var, value="emitidas").pack(side="left", padx=(0, 8))
        ttk.Radiobutton(format_frame, text="Recebidas", variable=self.download_pdf_tipo_var, value="recebidas").pack(side="left")

        toolbar = ttk.Frame(form)
        toolbar.grid(row=4, column=0, columnspan=10, sticky="ew", pady=(10, 0))

        ttk.Button(toolbar, text="Adicionar", command=self._add_row).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Atualizar selecionada", command=self._update_selected_row).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Aplicar datas aos selecionados", command=self._apply_dates_to_selected_rows).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Remover selecionadas", command=self._remove_selected_rows).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Limpar campos", command=self._clear_form).pack(side="left", padx=(0, 8))

        main = ttk.PanedWindow(self.root, orient="vertical")
        main.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))

        table_frame = ttk.Frame(main, padding=(0, 0, 0, 8))
        log_frame = ttk.Frame(main)
        main.add(table_frame, weight=2)
        main.add(log_frame, weight=1)

        self._build_table(table_frame)
        self._build_log(log_frame)

        bottom = ttk.Frame(self.root, padding=(16, 0, 16, 16))
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        left_actions = ttk.Frame(bottom)
        left_actions.grid(row=0, column=0, sticky="w")

        ttk.Button(left_actions, text="Salvar Excel", command=self._save_excel).pack(side="left", padx=(0, 8))
        ttk.Button(left_actions, text="Recarregar Excel", command=self._load_excel_if_exists).pack(side="left", padx=(0, 8))
        ttk.Button(left_actions, text="Abrir pasta notasfiscais", command=self._open_output_folder).pack(side="left")
        ttk.Button(left_actions, text="Gerar relatorio", command=self._export_report).pack(side="left", padx=(8, 0))

        right_actions = ttk.Frame(bottom)
        right_actions.grid(row=0, column=1, sticky="e", padx=(8, 0))

        self.only_selected_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(right_actions, text="Somente selecionadas", variable=self.only_selected_var).pack(side="left", padx=(0, 12))

        self.run_btn = ttk.Button(right_actions, text="Rodar busca", command=self._run_script)
        self.run_btn.pack(side="left", padx=(0, 8))

        self.stop_btn = ttk.Button(right_actions, text="Parar", command=self._stop_script, state="disabled")
        self.stop_btn.pack(side="left")

        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=220)
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def _add_field(self, parent, label, var, col, show=None):
        ttk.Label(parent, text=label).grid(row=0, column=col, columnspan=2, sticky="w", padx=(0, 8))
        entry = ttk.Entry(parent, textvariable=var, show=show if show else "")
        entry.grid(row=1, column=col, columnspan=2, sticky="ew", padx=(0, 8), pady=(2, 0))

    def _build_table(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(parent, columns=COLUMNS, show="headings", selectmode="extended")
        for col in COLUMNS:
            self.tree.heading(col, text=col)

        self.tree.column("Nome empresa", width=320, anchor="w")
        self.tree.column("cnpj", width=140, anchor="center")
        self.tree.column("senha", width=120, anchor="center")
        self.tree.column("data_inicio", width=120, anchor="center")
        self.tree.column("data_fim", width=120, anchor="center")

        yscroll = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<<TreeviewSelect>>", self._on_select_row)

    def _build_log(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = ttk.Frame(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top.columnconfigure(0, weight=1)

        ttk.Label(top, text="Log de processamento", style="Sub.TLabel").grid(row=0, column=0, sticky="w")

        self.status_var = tk.StringVar(value="Pronto")
        ttk.Label(top, textvariable=self.status_var).grid(row=0, column=1, sticky="e")

        self.log_text = tk.Text(parent, height=10, wrap="word", bg="#0f1720", fg="#d7e2ee", font=("Consolas", 10))
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

    def _append_log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_status(self, text):
        self.status_var.set(text)

    def _clear_form(self):
        self.nome_var.set("")
        self.cnpj_var.set("")
        self.senha_var.set("")
        self.data_inicio_var.set("")
        self.data_fim_var.set("")

    def _validate_date(self, value):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                datetime.strptime(value, fmt)
                return True
            except ValueError:
                pass
        return False

    def _get_form_data(self, require_identity=True):
        row = {
            "Nome empresa": self.nome_var.get().strip(),
            "cnpj": self.cnpj_var.get().strip(),
            "senha": self.senha_var.get().strip(),
            "data_inicio": self.data_inicio_var.get().strip(),
            "data_fim": self.data_fim_var.get().strip(),
        }

        if require_identity and (not row["Nome empresa"] or not row["cnpj"] or not row["senha"]):
            raise ValueError("Preencha Nome empresa, CNPJ e Senha.")

        if not self._validate_date(row["data_inicio"]):
            raise ValueError("data_inicio invalida. Use DD/MM/YYYY ou YYYY-MM-DD.")

        if not self._validate_date(row["data_fim"]):
            raise ValueError("data_fim invalida. Use DD/MM/YYYY ou YYYY-MM-DD.")

        return row

    def _add_row(self):
        try:
            row = self._get_form_data()
        except ValueError as e:
            messagebox.showerror("Dados invalidos", str(e))
            return

        self.tree.insert("", "end", values=[row[c] for c in COLUMNS])
        self._clear_form()

    def _update_selected_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Selecione", "Selecione uma linha para atualizar.")
            return

        try:
            row = self._get_form_data()
        except ValueError as e:
            messagebox.showerror("Dados invalidos", str(e))
            return

        self.tree.item(selected[0], values=[row[c] for c in COLUMNS])

    def _apply_dates_to_selected_rows(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Selecione", "Selecione uma ou mais linhas para aplicar as datas.")
            return

        try:
            row = self._get_form_data(require_identity=False)
        except ValueError as e:
            messagebox.showerror("Dados invalidos", str(e))
            return

        selected_set = set(selected)
        for item in self.tree.get_children():
            if item not in selected_set:
                continue
            values = list(self.tree.item(item, "values"))
            if len(values) != len(COLUMNS):
                continue
            values[3] = row["data_inicio"]
            values[4] = row["data_fim"]
            self.tree.item(item, values=values)

        messagebox.showinfo("Atualizado", f"Datas aplicadas a {len(selected)} empresa(s).")

    def _remove_selected_rows(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Selecione", "Selecione uma ou mais linhas para remover.")
            return

        for item in selected:
            self.tree.delete(item)

    def _get_selected_row_indexes(self):
        selected = set(self.tree.selection())
        if not selected:
            return []

        indexes = []
        for index, item in enumerate(self.tree.get_children()):
            if item in selected:
                indexes.append(index)
        return indexes

    def _on_select_row(self, _event=None):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0], "values")
        if len(values) != len(COLUMNS):
            return

        self.nome_var.set(values[0])
        self.cnpj_var.set(values[1])
        self.senha_var.set(values[2])
        self.data_inicio_var.set(values[3])
        self.data_fim_var.set(values[4])

    def _collect_rows(self):
        rows = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            rows.append(dict(zip(COLUMNS, values)))
        return rows

    def _save_excel(self):
        rows = self._collect_rows()
        if not rows:
            messagebox.showwarning("Sem dados", "Adicione pelo menos uma empresa antes de salvar.")
            return

        df = pd.DataFrame(rows, columns=COLUMNS)
        df.to_excel(EXCEL_PATH, index=False)
        self._append_log(f"[INFO] Excel salvo em: {EXCEL_PATH}\n")
        self._set_status("Excel salvo")

    def _load_excel_if_exists(self):
        self.tree.delete(*self.tree.get_children())

        if not EXCEL_PATH.exists():
            self._append_log("[INFO] empresas.xlsx ainda nao existe.\n")
            self._set_status("Aguardando cadastro")
            return

        try:
            df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
        except Exception as e:
            messagebox.showerror("Erro ao abrir Excel", str(e))
            return

        missing = [c for c in COLUMNS if c not in df.columns]
        if missing:
            messagebox.showerror("Colunas ausentes", f"Colunas ausentes no Excel: {missing}")
            return

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[row[c] for c in COLUMNS])

        self._append_log(f"[INFO] Excel carregado: {EXCEL_PATH}\n")
        self._set_status("Excel carregado")

    def _open_output_folder(self):
        folder = OUTPUT_DIR
        folder.mkdir(exist_ok=True)
        os.startfile(str(folder))

    def _start_progress(self):
        self.progress_running = True
        self.progress.start(10)

    def _stop_progress(self):
        self.progress_running = False
        self.progress.stop()

    def _stop_script(self):
        if not self.process or self.process.poll() is not None:
            return

        try:
            self.process.terminate()
            self._append_log("\n[INFO] Solicitado encerramento da execucao...\n")
            self._set_status("Encerrando...")
        except Exception as e:
            self._append_log(f"\n[ERRO] Nao foi possivel parar o processo: {e}\n")

    def _build_report_dataframe(self):
        if not OUTPUT_DIR.exists():
            return pd.DataFrame(columns=["empresa", "mes_ano", "tipo", "quantidade_xml"])

        registros = []
        for path in OUTPUT_DIR.rglob("*.xml"):
            try:
                rel = path.relative_to(OUTPUT_DIR)
                partes = rel.parts
                if len(partes) < 4:
                    continue
                empresa = partes[0]
                mes_ano = partes[1]
                tipo = partes[2]
                registros.append((empresa, mes_ano, tipo, 1))
            except Exception:
                continue

        if not registros:
            return pd.DataFrame(columns=["empresa", "mes_ano", "tipo", "quantidade_xml"])

        df = pd.DataFrame(registros, columns=["empresa", "mes_ano", "tipo", "qtd"])
        resumo = (
            df.groupby(["empresa", "mes_ano", "tipo"], as_index=False)["qtd"]
            .sum()
            .rename(columns={"qtd": "quantidade_xml"})
            .sort_values(["empresa", "mes_ano", "tipo"])
        )
        return resumo

    def _export_report(self):
        resumo = self._build_report_dataframe()
        if resumo.empty:
            messagebox.showinfo("Sem dados", "Nenhum XML encontrado para gerar relatorio.")
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = APP_DIR / f"relatorio_execucao_{stamp}.xlsx"
        resumo.to_excel(report_path, index=False)
        self._append_log(f"[INFO] Relatorio gerado: {report_path}\n")
        self._set_status("Relatorio gerado")
        messagebox.showinfo("Relatorio", f"Relatorio salvo em:\n{report_path}")

    def _run_script(self):
        if self.is_running:
            messagebox.showinfo("Em execucao", "A busca ja esta em execucao.")
            return

        if not SCRIPT_PATH.exists():
            if not getattr(sys, "frozen", False):
                messagebox.showerror("Script nao encontrado", f"Arquivo nao encontrado: {SCRIPT_PATH}")
                return

        rows = self._collect_rows()
        if not rows:
            messagebox.showwarning("Sem dados", "Adicione empresas antes de executar.")
            return

        if self.only_selected_var.get():
            self.selected_row_indexes = self._get_selected_row_indexes()
            if not self.selected_row_indexes:
                messagebox.showwarning("Selecione", "Marque 'Somente selecionadas' e escolha uma ou mais empresas.")
                return
        else:
            self.selected_row_indexes = None

        self._save_excel()

        self.is_running = True
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._start_progress()
        self._set_status("Processando...")
        self._append_log("\n[INFO] Iniciando processamento...\n")
        self._append_log(
            f"[INFO] Formatos selecionados: XML={'sim' if self.download_xml_var.get() else 'nao'} | "
            f"PDF={'sim' if self.download_pdf_var.get() else 'nao'} | "
            f"Tipo PDF={self.download_pdf_tipo_var.get()}\n"
        )
        if self.selected_row_indexes is not None:
            self._append_log(f"[INFO] Processando somente as linhas: {', '.join(map(str, self.selected_row_indexes))}\n")

        thread = threading.Thread(target=self._run_script_thread, daemon=True)
        thread.start()

    def _run_script_thread(self):
        if getattr(sys, "frozen", False):
            cmd = [str(Path(sys.executable).resolve()), "--worker"]
        else:
            cmd = [sys.executable, str(WORKER_SCRIPT_PATH)]

        # Prepara variáveis de ambiente para passar parâmetros
        env = os.environ.copy()
        env["NOVA_BUSCA_BAIXAR_XML"] = "1" if self.download_xml_var.get() else "0"
        env["NOVA_BUSCA_BAIXAR_PDF"] = "1" if self.download_pdf_var.get() else "0"
        env["NOVA_BUSCA_PDF_TIPO"] = self.download_pdf_tipo_var.get()
        if self.selected_row_indexes is not None:
            env["NOVA_BUSCA_LINHAS_SELECIONADAS"] = ",".join(str(i) for i in self.selected_row_indexes)

        try:
            self.process = subprocess.Popen(
                cmd,
                cwd=str(APP_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                env=env,
            )

            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.root.after(0, self._append_log, line)

            code = self.process.wait()
            if code == 0:
                self.root.after(0, self._set_status, "Concluido")
                self.root.after(0, self._append_log, "\n[OK] Processamento finalizado com sucesso.\n")
                self.root.after(0, self._append_log, "[INFO] Gere o relatorio para ver o resumo por empresa/periodo.\n")
            else:
                self.root.after(0, self._set_status, "Finalizado com erro")
                self.root.after(0, self._append_log, f"\n[ERRO] Script finalizou com codigo {code}.\n")

        except Exception as e:
            self.root.after(0, self._set_status, "Erro")
            self.root.after(0, self._append_log, f"\n[ERRO] Falha ao executar script: {e}\n")
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.run_btn.configure(state="normal"))
            self.root.after(0, lambda: self.stop_btn.configure(state="disabled"))
            self.root.after(0, self._stop_progress)
            self.selected_row_indexes = None


def run_worker_mode():
    buffer = io.StringIO()
    exit_code = 0
    baixar_xml = os.environ.get("NOVA_BUSCA_BAIXAR_XML", "1") == "1"
    baixar_pdf = os.environ.get("NOVA_BUSCA_BAIXAR_PDF", "0") == "1"
    pdf_tipo = os.environ.get("NOVA_BUSCA_PDF_TIPO", "todos")
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
        try:
            worker_main(baixar_xml=baixar_xml, baixar_pdf=baixar_pdf, pdf_tipo=pdf_tipo)
        except Exception as e:
            print(f"[ERRO] Falha fatal na execucao: {e}")
            exit_code = 1

    output = buffer.getvalue()
    if output:
        print(output, end="")
    raise SystemExit(exit_code)


def main():
    if "--worker" in sys.argv:
        run_worker_mode()

    root = tk.Tk()
    NovaBuscaUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
