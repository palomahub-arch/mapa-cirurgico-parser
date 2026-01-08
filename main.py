from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import shutil
import os
from datetime import datetime


from mapa_cirurgico import processar_lista_pdfs


app = FastAPI(title="Mapa Cirúrgico Web")


UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.get("/", response_class=HTMLResponse)
def home():
with open("templates/index.html", "r", encoding="utf-8") as f:
return f.read()


@app.post("/processar")
def processar_pdfs(pdfs: list[UploadFile] = File(...)):
caminhos = []


for pdf in pdfs:
caminho = os.path.join(UPLOAD_DIR, pdf.filename)
with open(caminho, "wb") as buffer:
shutil.copyfileobj(pdf.file, buffer)
caminhos.append(caminho)


data = datetime.now().strftime("%Y_%m_%d")
nome_saida = f"Mapa_Cirurgico_{data}.xlsx"
caminho_saida = os.path.join(OUTPUT_DIR, nome_saida)


caminho_final, total = processar_lista_pdfs(caminhos, OUTPUT_DIR)


return FileResponse(
path=caminho_final,
filename=nome_saida,
media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
