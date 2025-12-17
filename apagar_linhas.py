from openpyxl import load_workbook

def apagar_linhas_1_a_5(arquivo_entrada, arquivo_saida):
    wb = load_workbook(arquivo_entrada)
    ws = wb.worksheets[0]  # primeira aba

    # Apaga linhas 1 a 5
    ws.delete_rows(1, 5)

    wb.save(arquivo_saida)


# EXEMPLO DE USO
apagar_linhas_1_a_5(
    "entrada.xlsx",
    "saida_sem_linhas_1_a_5.xlsx"
)