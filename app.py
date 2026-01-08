import streamlit as st
from excel.limpar_excel import apagar_linhas
from pdf.limpar_pdf import pdf_utils

st.title("Limpeza de Arquivos")

tipo = st.radio("O que você quer limpar?", ["Excel", "PDF"])

arquivo = st.file_uploader(
    "Faça upload do arquivo",
    type=["xlsx", "xls"] if tipo == "Excel" else ["pdf"]
)

if arquivo:
    if tipo == "Excel":
        resultado = processar_excel(arquivo)
    else:
        resultado = processar_pdf(arquivo)

    st.success("Arquivo processado com sucesso!")
    st.download_button(
        "Baixar arquivo limpo",
        resultado,
        file_name=f"arquivo_limpo.{ 'xlsx' if tipo == 'Excel' else 'pdf' }"
    )

