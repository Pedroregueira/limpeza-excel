import streamlit as st
from openpyxl import load_workbook
import os

st.set_page_config(page_title="Teste - Apagar Linhas", layout="centered")
st.title("Apagar linhas 1 a 5")

def apagar_linhas_1_a_5(uploaded_file, arquivo_saida):
    wb = load_workbook(uploaded_file)
    ws = wb.worksheets[0]  # primeira aba

    # Apaga fisicamente as linhas 1 a 5
    ws.delete_rows(1, 5)

    wb.save(arquivo_saida)

arquivo = st.file_uploader("Envie o Excel", type=["xlsx"])

if arquivo:
    saida = "sem_linhas_1_a_5.xlsx"

    apagar_linhas_1_a_5(arquivo, saida)

    with open(saida, "rb") as f:
        st.download_button(
            "Baixar arquivo",
            f,
            file_name="sem_linhas_1_a_5.xlsx"
        )

    os.remove(saida)
