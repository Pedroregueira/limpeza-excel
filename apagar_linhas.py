import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
import os
import json
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Limpeza de Excel", layout="centered")
st.title("Limpeza automática de despesas")


# =========================
# FUNÇÃO PRINCIPAL DE LIMPEZA
# =========================
def limpar_excel(uploaded_file, arquivo_saida):
    wb = load_workbook(uploaded_file)

    aba_original = wb.worksheets[0]
    copia_original = wb.copy_worksheet(aba_original)
    copia_original.title = "ORIGINAL"

    ws = wb.worksheets[0]

    # 1) remover mesclagens
    merged_ranges = list(ws.merged_cells.ranges)
    for merged in merged_ranges:
        valor = ws.cell(row=merged.min_row, column=merged.min_col).value
        ws.unmerge_cells(str(merged))
        ws.cell(row=merged.min_row, column=merged.min_col).value = valor

    # 2) apagar linhas 1 a 5
    ws.delete_rows(1, 5)

    # 3) subir coluna I
    col_compl = 9
    ultima_linha = ws.max_row
    for row in range(3, ultima_linha + 1):
        ws.cell(row=row - 1, column=col_compl).value = ws.cell(row=row, column=col_compl).value
    ws.cell(row=ultima_linha, column=col_compl).value = None

    # 4) apagar linhas com data vazia (coluna D)
    col_data = 4
    for row in range(ws.max_row, 1, -1):
        if ws.cell(row=row, column=col_data).value in (None, ""):
            ws.delete_rows(row)

    wb.save(arquivo_saida)


# =========================
# APPEND NO GOOGLE SHEETS
# =========================
def append_to_google_sheets(nome_arquivo_excel, spreadsheet_id, aba_nome):

    creds = Credentials.from_service_account_info(
        json.loads(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )

    gc = gspread.authorize(creds)

    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(aba_nome)

    wb = load_workbook(nome_arquivo_excel)
    ws_excel = wb.active

    dados = [
        [cell.value for cell in row]
        for row in ws_excel.iter_rows()
    ]

    # adiciona após a última linha preenchida
    ws.append_rows(dados, value_input_option="USER_ENTERED")

    return sh.url


# =========================
# INTERFACE STREAMLIT
# =========================
arquivo = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

if arquivo:

    data_atual = datetime.now()
    mes = data_atual.strftime("%m")
    ano = data_atual.strftime("%Y")

    nome_saida = f"RLT_DESPESAS_{mes}_{ano}.xlsx"

    limpar_excel(arquivo, nome_saida)

    st.success("Arquivo tratado com sucesso ✅")

    # >>> INFORME AQUI <<<
    SPREADSHEET_ID = "/d/163fTHvX6-ygJD0RPnOwAmbKspUnE4olv7JoX_dUcJJk/edit?gid=0#gid=0"
    ABA_DESTINO = "RLTDESPESAS"

    url = append_to_google_sheets(
        nome_arquivo_excel=nome_saida,
        spreadsheet_id=SPREADSHEET_ID,
        aba_nome=ABA_DESTINO
    )

    st.success("Dados enviados para o Google Sheets 🚀")
    st.write(url)

    os.remove(nome_saida)

