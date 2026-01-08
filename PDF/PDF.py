import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# ==============================
# CONFIG STREAMLIT (PADRÃO CLIENTE)
# ==============================
st.set_page_config(
    page_title="Extrator de Horas (PDF)",
    layout="wide"
)

# Centralização visual
st.markdown(
    """
    <style>
        .block-container {
            max-width: 1200px;
            margin: auto;
            padding-top: 2rem;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("## 📄 Extrator de Horas (PDF)")
st.markdown(
    "Aplicação para extração automática de horas a partir de arquivos PDF, "
    "com geração de relatório estruturado."
)
st.divider()

# ==============================
# UPLOAD
# ==============================
uploaded_pdf = st.file_uploader(
    "📎 Envie o arquivo PDF",
    type=["pdf"]
)

# ==============================
# FUNÇÃO DE EXTRAÇÃO
# ==============================
def extrair_pdf(uploaded_pdf):
    linhas = []

    padrao = re.compile(r"^(.*?)\s+(\d+:\d{2})\s+([\d\.]+,\d{2})$")

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue

            for linha in texto.split("\n"):
                linha = linha.strip()

                if (
                    linha.startswith("Reporte de Horas")
                    or linha.startswith("Período")
                    or linha.startswith("Empresa:")
                    or linha.startswith("Emitido em")
                    or linha.startswith("Total GERAL")
                    or linha == "Projeto Horas Valor"
                ):
                    continue

                match = padrao.match(linha)
                if match:
                    projeto, horas, valor = match.groups()
                    linhas.append({
                        "Projeto": projeto,
                        "Horas": horas,
                        "Valor": valor
                    })

    return pd.DataFrame(linhas)

# ==============================
# PROCESSAMENTO
# ==============================
if uploaded_pdf:
    st.success("PDF carregado com sucesso!")

    if st.button("🔍 Extrair dados"):
        df = extrair_pdf(uploaded_pdf)

        if df.empty:
            st.warning("Nenhum dado encontrado no PDF.")
        else:
            st.markdown("### 📑 Pré-visualização do relatório")

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

            st.divider()

            buffer = BytesIO()
            df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                label="⬇️ Baixar relatório em Excel",
                data=buffer,
                file_name="RLT_HORAS_PDF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
