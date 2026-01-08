import logging
from logger_config import configurar_logger
configurar_logger()

import sys
sys.dont_write_bytecode = True

import tkinter as tk
from tkinter import filedialog, messagebox
import os

from mapa_cirurgico import processar_lista_pdfs


class MapaCirurgicoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mapa Cirúrgico - PDF para Excel")
        self.root.geometry("500x270")
        self.root.resizable(False, False)

        self.lista_pdfs = []

        # ===============================
        # TÍTULO / INSTRUÇÃO
        # ===============================
        self.label_info = tk.Label(
            root,
            text="Selecione os arquivos PDF do mapa cirúrgico",
            font=("Arial", 11)
        )
        self.label_info.pack(pady=15)

        # ===============================
        # BOTÃO SELEÇÃO
        # ===============================
        self.btn_selecionar = tk.Button(
            root,
            text="Selecionar PDFs",
            width=25,
            command=self.selecionar_pdfs
        )
        self.btn_selecionar.pack(pady=5)

        # ===============================
        # QUANTIDADE DE ARQUIVOS
        # ===============================
        self.label_qtd = tk.Label(
            root,
            text="Nenhum arquivo selecionado"
        )
        self.label_qtd.pack(pady=5)

        # ===============================
        # BOTÃO PROCESSAR
        # ===============================
        self.btn_processar = tk.Button(
            root,
            text="Processar",
            width=25,
            state=tk.DISABLED,
            command=self.processar
        )
        self.btn_processar.pack(pady=10)

        # ===============================
        # STATUS (RODAPÉ)
        # ===============================
        self.status_var = tk.StringVar(value="🟡 Aguardando seleção de arquivos")
        self.label_status = tk.Label(
            root,
            textvariable=self.status_var,
            anchor="w",
            relief=tk.SUNKEN,
            padx=10
        )
        self.label_status.pack(side=tk.BOTTOM, fill=tk.X)

    # =====================================================
    # SELEÇÃO DE PDFs
    # =====================================================
    def selecionar_pdfs(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecione os PDFs",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )

        if arquivos:
            self.lista_pdfs = list(arquivos)
            self.label_qtd.config(
                text=f"{len(self.lista_pdfs)} arquivo(s) selecionado(s)"
            )
            self.btn_processar.config(state=tk.NORMAL)
            self.status_var.set("🟢 Arquivos selecionados com sucesso")

    # =====================================================
    # PROCESSAMENTO
    # =====================================================
    def processar(self):
        if not self.lista_pdfs:
            messagebox.showwarning(
                "Atenção",
                "Nenhum arquivo PDF foi selecionado."
            )
            return

        try:
            # UI - estado processando
            self.status_var.set("🔵 Processando arquivos, aguarde...")
            self.btn_processar.config(state=tk.DISABLED)
            self.btn_selecionar.config(state=tk.DISABLED)
            self.root.update_idletasks()

            pasta_saida = os.path.join(os.path.expanduser("~"), "Downloads")

            caminho, total = processar_lista_pdfs(
                self.lista_pdfs,
                pasta_saida
            )

            self.status_var.set("🟢 Processamento concluído com sucesso")

            messagebox.showinfo(
                "Sucesso",
                f"Processamento concluído!\n\n"
                f"Registros gerados: {total}\n"
                f"Arquivo salvo em:\n{caminho}"
            )

        except Exception as e:
            logging.exception("Erro durante o processamento")
            self.status_var.set("🔴 Erro durante o processamento")

            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro durante o processamento:\n\n{str(e)}"
            )

        finally:
            # UI - volta ao normal
            self.btn_selecionar.config(state=tk.NORMAL)
            self.btn_processar.config(
                state=tk.NORMAL if self.lista_pdfs else tk.DISABLED
            )


# =====================================================
# EXECUÇÃO
# =====================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = MapaCirurgicoApp(root)
    root.mainloop()
