import pdfplumber
import pandas as pd
import re
import os
import logging
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

logging.getLogger("pdfminer").setLevel(logging.ERROR)

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================

def extrai_data(pagina):
    texto = pagina.extract_text() or ""
    m = re.search(r'Data:\s*(\d{2}/\d{2}/\d{4})', texto)
    return m.group(1) if m else None


def heuristica_palavras_quebradas(texto):
    if not isinstance(texto, str):
        return texto

    correcoes = {
        r'INTELIGE\s*NTE': 'INTELIGENTE',
        r'ROBOTIC\s*A': 'ROBOTICA',
        r'ROBÓTIC\s*A': 'ROBÓTICA',
        r'HEMODIN\s*AMICA': 'HEMODINAMICA'
    }

    for p, c in correcoes.items():
        texto = re.sub(p, c, texto, flags=re.I)

    return re.sub(r'\s+', ' ', texto).strip()


def unidades_tratamento(texto):
    if not isinstance(texto, str):
        return texto

    regras = {
        r'^Nova\s*Lima.*': 'Unidade Nova Lima',
        r'^Contorno.*': 'Unidade Contorno',
        r'^Betim.*': 'Unidade Betim Contagem',
        r'^Emec.*': 'EMEC',
        r'^Premium.*': 'Hospital Mater Dei Premium'
    }

    for p, c in regras.items():
        texto = re.sub(p, c, texto, flags=re.I)

    return texto.strip()

# =====================================================
# PROCESSAMENTO DE PÁGINA
# =====================================================

def processar_pagina(pagina):
    tabelas = pagina.extract_tables()
    df = pd.DataFrame()

    if not tabelas:
        return df

    linhas = []
    for tabela in tabelas:
        for linha in tabela:
            linhas.append([
                c.replace('\n', ' ').strip() if isinstance(c, str) else ''
                for c in linha
            ])

    # -------------------------------------------------
    # DETECÇÃO DO CABEÇALHO
    # -------------------------------------------------
    idx = None
    for i, l in enumerate(linhas):
        t = " ".join(l).upper()
        if 'SALA' in t and 'HORA' in t and 'PROCEDIMENTO' in t:
            if not any(re.match(r'\d{2}:\d{2}', c) for c in l):
                idx = i
                break

    if idx is None:
        return df

    cabecalho = linhas[idx]
    dados = linhas[idx + 1:]

    # -------------------------------------------------
    # MERGE DE LINHAS QUEBRADAS
    # -------------------------------------------------
    dados_mesclados = []
    linha_pendente = None

    for linha in dados:
        tem_horario = any(re.match(r'\d{2}:\d{2}', c) for c in linha[:5])

        if tem_horario:
            if linha_pendente:
                dados_mesclados.append(linha_pendente)
            linha_pendente = linha
        elif linha_pendente:
            palavras_cabecalho = ['HORA', 'SALA', 'AVISO', 'NOME', 'PROCEDIMENTO', 'DR.PRE']
            for i, v in enumerate(linha):
                if i < len(linha_pendente) and v:
                    if any(palavra in v.upper() for palavra in palavras_cabecalho):
                        continue
                    
                    if not linha_pendente[i]:
                        linha_pendente[i] = v
                    elif len(v) > 5:
                        linha_pendente[i] = f"{linha_pendente[i]} {v}".strip()
        else:
            if any(linha):
                dados_mesclados.append(linha)

    if linha_pendente:
        dados_mesclados.append(linha_pendente)

    # -------------------------------------------------
    # NORMALIZAÇÃO DO TAMANHO DAS LINHAS
    # -------------------------------------------------
    num_cols = len(cabecalho)
    linhas_corrigidas = []

    for linha in dados_mesclados:
        if len(linha) > num_cols:
            linha = linha[:num_cols - 1] + [' '.join(linha[num_cols - 1:])]
        elif len(linha) < num_cols:
            linha = linha + [''] * (num_cols - len(linha))
        linhas_corrigidas.append(linha)

    df = pd.DataFrame(linhas_corrigidas, columns=cabecalho)

    # -------------------------------------------------
    # RENAME DE COLUNAS
    # -------------------------------------------------
    df = df.rename(columns={
        'SALA': 'Local',
        'Sala': 'Local',
        'PROCEDIMENTO': 'Subatividade',
        'Procedimento:': 'Subatividade',
        'Procedimento': 'Subatividade',
        'HORA': 'Hora inicio',
        'DR.PRE V': 'Duração (min)',
        'DR.PREV': 'Duração (min)',
        'DRPREV': 'Duração (min)',
        'ANESTESISTAS': 'Profissional (GH)',
        'Anestesistas': 'Profissional (GH)',
        'CIRURGIAO': 'Agente externo',
        'Cirurgiao': 'Agente externo'
        })

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna('').str.strip()

    df['Local'] = (
        df['Local']
        .replace('', pd.NA)
        .ffill()
        .apply(heuristica_palavras_quebradas)
    )

    df = df[df['Hora inicio'].str.match(r'\d{2}:\d{2}', na=False)]

    # -------------------------------------------------
    # LIMPEZA DA COLUNA HORA
    # -------------------------------------------------
    if 'Hora inicio' in df.columns:
        df['Hora inicio'] = df['Hora inicio'].apply(
            lambda x: re.search(r'\d{2}:\d{2}', str(x)).group() 
            if re.search(r'\d{2}:\d{2}', str(x)) else x
        )

    df['Duração (min)'] = pd.to_datetime(
        df['Duração (min)'], format='%H:%M', errors='coerce'
    )
    df['Duração (min)'] = (
        df['Duração (min)'].dt.hour * 60 +
        df['Duração (min)'].dt.minute
    )

    # -------------------------------------------------
    # LIMPEZA DE PALAVRAS-LIXO
    # -------------------------------------------------
    palavras_lixo = [
        'CIRURGIAO', 'ANESTESISTAS', 'EQUIPAMENTOS', 'MATERIAIS',
        'ATEND.', 'ATEND', 'OBSERVAÇÕES', 'OBSERVACOES'
    ]
    
    colunas_limpar = [
        'Subatividade', 'Agente externo', 'Profissional (GH)'
    ]
    
    for col in colunas_limpar:
        if col in df.columns:
            for palavra in palavras_lixo:
                padrao_fim = r'\s*' + palavra + r'\s*$'
                padrao_inicio = r'^\s*' + palavra + r'\s*'
                df[col] = df[col].str.replace(padrao_fim, '', regex=True, case=False)
                df[col] = df[col].str.replace(padrao_inicio, '', regex=True, case=False)
            df[col] = df[col].str.strip()

    # -------------------------------------------------
    # NORMALIZAÇÃO DA SALA ROBÓTICA
    # -------------------------------------------------
    df['Local'] = df['Local'].apply(
        lambda x: 'SALA ROBOTICA (NL)'
        if isinstance(x, str) and 'ROBOTIC' in x.upper()
        else x
    )

    # -------------------------------------------------
    # LIMPEZA CAMPOS DA SALA ROBÓTICA
    # -------------------------------------------------
    if 'Local' in df.columns:
        mask_robotica = df['Local'].str.contains('ROBOTICA', case=False, na=False)
        if mask_robotica.any():
            campos_limpar = [
                'Subatividade', 'Duração (min)', 'Profissional (GH)',
                'Agente externo', 'Paciente', 'Aviso Cirurgico'
            ]
            for campo in campos_limpar:
                if campo in df.columns:
                    df.loc[mask_robotica, campo] = ''

    colunas_finais = [
        'Local', 'Subatividade', 'Hora inicio', 'Duração (min)',
        'Profissional (GH)', 'Agente externo'
    ]

    for col in colunas_finais:
        if col not in df.columns:
            df[col] = ''

    return df[colunas_finais]

# =====================================================
# PROCESSAMENTO DA PASTA
# =====================================================

def processar_pasta(pasta_pdfs):
    arquivos = [
        os.path.join(pasta_pdfs, f)
        for f in os.listdir(pasta_pdfs)
        if f.lower().endswith('.pdf')
    ]

    df_final = pd.DataFrame()

    for pdf_path in arquivos:
        unidade = os.path.splitext(os.path.basename(pdf_path))[0]
        print(f"\n📄 Processando: {unidade}")

        with pdfplumber.open(pdf_path) as pdf:
            data_extraida = extrai_data(pdf.pages[0])

            for pagina in pdf.pages:
                df_aux = processar_pagina(pagina)
                if df_aux.empty:
                    continue

                df_aux['Data'] = data_extraida
                df_aux['Unidade'] = unidades_tratamento(unidade)
                df_aux['Escala'] = None

                df_final = pd.concat([df_final, df_aux], ignore_index=True)
    
    colunas_proibidas = ['Paciente', 'Aviso Cirurgico']
    df_final = df_final.drop(
        columns=[c for c in colunas_proibidas if c in df_final.columns],
        errors='ignore'
    )

    # =================================================
    # EXPORTAÇÃO EXCEL
    # =================================================
    wb = Workbook()
    ws = wb.active
    ws.title = "Mapa de Atividades"

    ws.append(['Obrigatorio'] * 4 + ['Opcional'] * 5)

    colunas = [
        'Data', 'Unidade', 'Escala', 'Local', 'Subatividade',
        'Hora inicio', 'Duração (min)', 'Profissional (GH)',
        'Agente externo'
    ]
    ws.append(colunas)

    for r in dataframe_to_rows(df_final[colunas], index=False, header=False):
        ws.append(r)

    caminho = os.path.join(pasta_pdfs, 'Mapa Cirurgico.xlsx')
    wb.save(caminho)

    print(f"\n✅ Excel gerado com sucesso: {caminho}")
    print(f"✅ Total de registros: {len(df_final)}")

# =====================================================
# EXECUÇÃO
# =====================================================

# pasta_com_pdfs = r"G:\Meu Drive\Projetos\8 - Mapa Cirúrgico\NF to xlsx"
# processar_pasta(pasta_com_pdfs)