import streamlit as st
from openpyxl import load_workbook
import os

st.set_page_config(page_title="Limpeza - Etapa 3", layout="centered")
st.title("Limpeza básica do Excel")

def limpar_excel(uploaded_file, arquivo_saida):
    wb = load_workbook(uploaded_file)
    ws = wb.worksheets[0]  # primeira aba

    # =========================
    # 1) Remover mesclagem (sem replicar)
    # =========================
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

    # =========================
    # 2) Apagar linhas 1 a 5
    # =========================
    ws.delete_rows(1, 5)

    # =========================
    # 3) Subir coluna I (Compl.lcto) 1 linha
    # =========================
    col_compl = 9  # coluna I
    ultima_linha = ws.max_row

    for row in range(2, ultima_linha + 1):
        ws.cell(row=row - 1, column=col_compl).value = (
            ws.cell(row=row, column=col_compl).value
        )

    ws.cell(row=ultima_linha, column=col_compl).value = None

    # =========================
    # 4) Apagar linhas com Dt.lançtos vazio (coluna D)
    # =========================
    col_data = 4  # coluna D

    for row in range(ws.max_row, 1, -1):
        if ws.cell(row=row, column=col_data).value in (None, ""):
            ws.delete_rows(row)

    wb.save(arquivo_saida)

arquivo = st.file_uploader("Envie o Excel", type=["xlsx"])

if arquivo:
    saida = "limpo_final.xlsx"

    limpar_excel(arquivo, saida)

    with open(saida, "rb") as f:
        st.download_button(
            "Baixar arquivo limpo",
            f,
            file_name="limpo_final.xlsx"
        )

    os.remove(saida)
