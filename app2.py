import streamlit as st

# CORREÇÃO AQUI: Importar a função correta
# Certifique-se que dentro de excel/limpar_excel.py existe uma função "processar_excel"
from excel.limpar_excel import processar_excel 
from pdf.limpar_pdf import processar_pdf

st.title("Limpeza de Arquivos")

tipo = st.radio("O que você quer limpar?", ["Excel", "PDF"])

# ... resto do código de upload ...
arquivo = st.file_uploader(
    "Faça upload do arquivo",
    type=["xlsx", "xls"] if tipo == "Excel" else ["pdf"]
)

if arquivo:
    if tipo == "Excel":
        # Agora esta função existe pois foi importada lá em cima
        resultado = processar_excel(arquivo) 
    else:
        resultado = processar_pdf(arquivo)
    
    st.success("Arquivo processado com sucesso!")
    # ... código de download ...
