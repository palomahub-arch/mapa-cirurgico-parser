import sys
sys.dont_write_bytecode = True

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from datetime import datetime

from mapa_cirurgico import processar_pasta


def selecionar_pdfs():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os PDFs do Mapa Cirúrgico",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if not arquivos:
        return

    # Pasta Downloads do usuário
    pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")

    # Pasta temporária para os PDFs
    pasta_temp = os.path.join(pasta_downloads, "pdfs_temp")

    try:
        os.makedirs(pasta_temp, exist_ok=True)

        # Copiar PDFs selecionados para a pasta temporária
        for arq in arquivos:
            shutil.copy(arq, pasta_temp)

        # Processar PDFs
        processar_pasta(pasta_temp)

        # Nome padronizado do Excel
        data_hoje = datetime.now().strftime("%Y_%m_%d")
        nome_excel = f"Mapa Cirúrgico {data_hoje}.xlsx"

        excel_temp = os.path.join(pasta_temp, "Mapa Cirurgico.xlsx")
        excel_final = os.path.join(pasta_downloads, nome_excel)

        # Mover Excel para Downloads
        if os.path.exists(excel_temp):
            shutil.move(excel_temp, excel_final)
        else:
            raise FileNotFoundError("Arquivo Excel não foi gerado.")

        messagebox.showinfo(
            "Sucesso",
            f"Mapa Cirúrgico gerado com sucesso!\n\n"
            f"Arquivo salvo em:\n{excel_final}"
        )

    except Exception as e:
        messagebox.showerror("Erro", str(e))

    finally:
        # Remove a pasta temporária
        shutil.rmtree(pasta_temp, ignore_errors=True)


# =========================
# INTERFACE GRÁFICA
# =========================

root = tk.Tk()
root.title("Gerador de Mapa Cirúrgico")
root.geometry("420x180")
root.resizable(False, False)

btn = tk.Button(
    root,
    text="Selecionar PDFs e Gerar Excel",
    font=("Arial", 12),
    width=32,
    height=3,
    command=selecionar_pdfs
)

btn.pack(expand=True)

root.mainloop()
