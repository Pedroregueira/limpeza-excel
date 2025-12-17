import streamlit as st
from openpyxl import load_workbook
import os

st.set_page_config(page_title="Teste - Remover Mesclagem", layout="centered")
st.title("Remover células mescladas")

def remover_mesclagem(uploaded_file, arquivo_saida):
    wb = load_workbook(uploaded_file)
    ws = wb.worksheets[0]  # primeira aba

    # Copiamos a lista porque ela muda durante o loop
    merged_ranges = list(ws.merged_cells.ranges)

    for merged in merged_ranges:
        # Valor da célula superior esquerda
        valor = ws.cell(
            row=merged.min_row,
            column=merged.min_col
        ).value

        # Desmescla
        ws.unmerge_cells(str(merged))

        # Preenche todas as células com o valor
        for row in ws.iter_rows(
            min_row=merged.min_row,
            max_row=merged.max_row,
            min_col=merged.min_col,
            max_col=merged.max_col
        ):
            for cell in row:
                cell.value = valor

    wb.save(arquivo_saida)

arquivo = st.file_uploader("Envie o Excel", type=["xlsx"])

if arquivo:
    saida = "sem_mesclagem.xlsx"

    remover_mesclagem(arquivo, saida)

    with open(saida, "rb") as f:
        st.download_button(
            "Baixar arquivo",
            f,
            file_name="sem_mesclagem.xlsx"
        )

    os.remove(saida)
