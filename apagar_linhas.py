import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
import os

st.set_page_config(page_title="Limpeza de Excel", layout="centered")
st.title("Limpeza automática de despesas")

def limpar_excel(uploaded_file, arquivo_saida):
    wb = load_workbook(uploaded_file)

    # =================================================
    # 0) Guardar aba original como página 2
    # =================================================
    aba_original = wb.worksheets[0]
    copia_original = wb.copy_worksheet(aba_original)
    copia_original.title = "ORIGINAL"

    ws = wb.worksheets[0]  # trabalhar apenas na primeira aba

    # =================================================
    # 1) Remover mesclagem (sem replicar valores)
    # =================================================
    merged_ranges = list(ws.merged_cells.ranges)

    for merged in merged_ranges:
        valor = ws.cell(
            row=merged.min_row,
            column=merged.min_col
        ).value

        ws.unmerge_cells(str(merged))
        ws.cell(
            row=merged.min_row,
            column=merged.min_col
        ).value = valor

    # =================================================
    # 2) Apagar linhas 1 a 5
    # =================================================
    ws.delete_rows(1, 5)

    # =================================================
    # 3) Subir coluna I (Compl.lcto) a partir da linha 2
    # =================================================
    col_compl = 9  # coluna I
    ultima_linha = ws.max_row

    for row in range(3, ultima_linha + 1):
        ws.cell(row=row - 1, column=col_compl).value = (
            ws.cell(row=row, column=col_compl).value
        )

    ws.cell(row=ultima_linha, column=col_compl).value = None

    # =================================================
    # 4) Remover linhas com Desc.cta EXACT
    #    "CUSTO C/ TERCEIROS PESSOA JURIDI"
    # =================================================
    col_desc_cta = 5  # coluna E

    for row in range(ws.max_row, 1, -1):
        valor = ws.cell(row=row, column=col_desc_cta).value
    if valor and valor.strip() == "CUSTO C/ TERCEIROS PESSOA JURIDI":
    ws.delete_rows(row)

    # =================================================
    # 5) Apagar linhas com Dt.lançtos vazio (coluna D)
    # =================================================
    col_data = 4  # coluna D

    for row in range(ws.max_row, 1, -1):
        if ws.cell(row=row, column=col_data).value in (None, ""):
            ws.delete_rows(row)

    wb.save(arquivo_saida)

# =================================================
# INTERFACE STREAMLIT
# =================================================
arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

if arquivo:
    data_atual = datetime.now()
    mes = data_atual.strftime("%m")
    ano = data_atual.strftime("%Y")

    nome_saida = f"RLT_DESPESAS_{mes}_{ano}.xlsx"

    limpar_excel(arquivo, nome_saida)

    with open(nome_saida, "rb") as f:
        st.download_button(
            label="Baixar relatório final",
            data=f,
            file_name=nome_saida
        )

    os.remove(nome_saida)
