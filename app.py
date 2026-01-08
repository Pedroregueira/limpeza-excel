import streamlit as st
import os
from excel_utils import limpar_excel
from pdf_utils import extrair_pdf

st.set_page_config(page_title="Ferramentas", layout="centered")
st.title("Ferramentas Financeiras")

tab1, tab2 = st.tabs(["📊 Excel", "📄 PDF"])

with tab1:
    excel = st.file_uploader("Envie o Excel", type=["xlsx"])
    if excel:
        nome = "EXCEL_LIMPO.xlsx"
        limpar_excel(excel, nome)
        with open(nome, "rb") as f:
            st.download_button("Baixar Excel", f, file_name=nome)
        os.remove(nome)

with tab2:
    pdf = st.file_uploader("Envie o PDF", type=["pdf"])
    if pdf:
        nome = "PDF_CONVERTIDO.xlsx"
        extrair_pdf(pdf, nome)
        with open(nome, "rb") as f:
            st.download_button("Baixar Excel", f, file_name=nome)
        os.remove(nome)
