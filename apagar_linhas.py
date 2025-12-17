import streamlit as st
from openpyxl import load_workbook
import os

st.set_page_config(page_title="Limpeza - Etapa 1", layout="centered")
st.title("Limpeza básica do Excel")

def limpar_excel(uploaded_file, arquivo_saida):
    wb = load_workbook(uploaded_file)
    ws = wb.worksheets[0]  # primeira aba

    # 1) Remover mesclagem
    merged_ranges = list(ws.merged_cells.ranges)

    for merged in merged_ranges:
        valor = ws.cell(
            row=merged.min_row,
            column=merged.min_col
        ).value

        ws.unmerge_cells(str(merged))

        for row in ws.iter_rows(
            min_row=merged.min_row,
            max_row=merged.max_row,
            min_col=merged.min_col,
            max_col=merged.max_col
        ):
            for cell in row:
                cell.value = valor

    # 2) Apagar linhas 1 a 5
    ws.delete_rows(1, 5)

    wb.save(arquivo_saida)

arquivo = st.file_uploader("Envie o Excel", type=["xlsx"])

if arquivo:
    saida = "limpo_etapa1.xlsx"

    limpar_excel(arquivo, saida)

    with open(saida, "rb") as f:
        st.download_button(
            "Baixar arquivo limpo",
            f,
            file_name="limpo_etapa1.xlsx"
        )

    os.remove(saida)
