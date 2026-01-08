# 🏥 Mapa Cirúrgico – PDF → Excel

Aplicação desenvolvida em **Python** para extração automática de mapas cirúrgicos em **PDF** e conversão para **Excel padronizado**, resolvendo problemas comuns de PDFs não estruturados (tabelas quebradas, linhas mescladas, cabeçalhos inconsistentes).

---

## 🚀 Visão Geral

Este projeto nasceu de uma necessidade real: transformar mapas cirúrgicos em PDF — frequentemente inconsistentes — em uma base estruturada e confiável para uso operacional e analítico.

A solução:

* Lê múltiplos PDFs
* Reconstrói tabelas quebradas
* Normaliza textos e unidades
* Remove campos sensíveis
* Exporta um Excel padronizado
* Pode ser distribuída como **executável (.exe)**

---

## 🧠 Principais Desafios Resolvidos

* PDFs sem padrão de estrutura
* Cabeçalhos variando por página
* Linhas de procedimentos quebradas
* Campos deslocados entre colunas
* Padronização de salas (ex.: robótica)
* Remoção definitiva de colunas sensíveis

---

## 🛠️ Tecnologias Utilizadas

* **Python 3**
* **pdfplumber** – leitura de PDFs
* **pandas** – tratamento e transformação de dados
* **openpyxl** – geração de Excel
* **PyInstaller** – empacotamento em executável

---

## 📁 Estrutura do Projeto

```
mapa-cirurgico-parser/
│
├── app.py                 # Ponto de entrada do aplicativo
├── mapa_cirurgico.py      # Lógica de extração e transformação
├── requirements.txt
├── README.md
├── examples/
│   ├── exemplo_entrada.pdf
│   └── exemplo_saida.xlsx

```

---

## ▶️ Como Executar (Modo Desenvolvimento)

```bash
pip install -r requirements.txt
python app.py
```

---

## 📦 Gerar Executável (.exe)

```powershell
python -m PyInstaller --onefile --noconsole --name mapa_cirurgico_app app.py
```

O executável final será gerado em:

```
dist/mapa_cirurgico_app.exe
```

---

## 📊 Estrutura do Excel Gerado

Colunas finais:

* Data
* Unidade
* Escala
* Local
* Subatividade
* Hora início
* Duração (min)
* Profissional (GH)
* Agente externo

> ⚠️ Campos sensíveis como paciente e aviso cirúrgico são removidos por design.

---

## 🧪 Exemplos

A pasta `examples/` contém PDFs e Excel **anonimizados**, apenas para demonstração.

---

## 🧭 Roadmap (Próximas Evoluções)

* [ ] Interface gráfica (desktop)
* [ ] Validação visual dos dados extraídos
* [ ] Sistema de logs
* [ ] Versão web (API)
* [ ] Testes automatizados

---

## 👤 Autor

Projeto desenvolvido por **Paloma Cristiane**

> Este projeto faz parte do meu portfólio profissional e reflete soluções aplicadas a problemas reais de dados.

---

## 📄 Licença

Uso educacional e demonstrativo.
