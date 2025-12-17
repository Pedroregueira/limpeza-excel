import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os

st.set_page_config(page_title="Limpeza de Excel", layout="centered")

def desmesclar_pagina_6(arquivo_entrada, arquivo_temp):
    wb = load_workbook(arquivo_entrada)
    ws = wb.worksheets[0]  # página 6 fixa

    for merged in list(ws.merged_cells.ranges):
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

    wb.save(arquivo_temp)

def tratar_excel(arquivo_entrada, arquivo_saida):
    arquivo_temp = "temp_sem_mescla.xlsx"

    desmesclar_pagina_6(arquivo_entrada, arquivo_temp)

    df = pd.read_excel(arquivo_temp, sheet_name=0)
    df_original = df.copy()

    df.iloc[:, 8] = df.iloc[:, 8].shift(-1)
    df = df[df["DT_LANÇTOS"].notna()]

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ARRUMADO", index=False)
        df_original.to_excel(writer, sheet_name="ORIGINAL", index=False)

    os.remove(arquivo_temp)

st.title("Limpeza de Excel")

arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

if arquivo:
    saida = "resultado.xlsx"
    tratar_excel(arquivo, saida)

    with open(saida, "rb") as f:
        st.download_button(
            "Baixar arquivo tratado",
            f,
            file_name="arquivo_tratado.xlsx"
        )

    os.remove(saida)

