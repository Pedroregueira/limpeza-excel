import pdfplumber
import pandas as pd
import re

def extrair_pdf(uploaded_pdf, arquivo_saida):
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

    pd.DataFrame(linhas).to_excel(arquivo_saida, index=False)
