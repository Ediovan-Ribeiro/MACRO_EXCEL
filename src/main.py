from openpyxl import load_workbook
import os
from datetime import datetime
import win32com.client as win32

# Caminhos
CAMINHO_ENTRADA = os.path.join('..', 'dados', 'entrada', 'exemplo.xlsm')
CAMINHO_SAIDA = os.path.join('..', 'dados', 'saida')

# Verificar e criar diretórios
if not os.path.exists(CAMINHO_ENTRADA):
    raise FileNotFoundError(f"Erro: Arquivo não encontrado: {CAMINHO_ENTRADA}")
os.makedirs(CAMINHO_SAIDA, exist_ok=True)

# Carregar o arquivo Excel
try:
    wb = load_workbook(CAMINHO_ENTRADA, keep_vba=True)
    aba_1of, aba_modelo = wb['1OF'], wb['MODELO']
except KeyError as e:
    raise KeyError(f"Erro: Aba não encontrada no arquivo: {e}")

# Mapeamento das colunas e linhas
MAPPING = {
    'A': (10, 1), 'B': (10, 2), 'C': (10, 3), 'D': (13, 1), 'E': (13, 2), 'F': (13, 3),
    'G': (16, 1), 'H': (16, 2), 'J': (18, 1), 'K': (21, 1), 'I': (23, 1), 'L': (23, 3),
    'O': (25, 1), 'P': (25, 2), 'N': (25, 3), 'M': (28, 1), 'Q': (28, 2), 'R': (28, 3),
    'S': (31, 1), 'T': (31, 2), 'U': (31, 3), 'V': (34, 1), 'W': (34, 2), 'X': (34, 3),
    'Y': (37, 1), 'Z': (37, 2), 'AA': (37, 3)
}

# Função para transferir dados
def transferir_dados(linha_1of, nova_aba_modelo):
    for col, (linha_dest, col_dest) in MAPPING.items():
        valor = aba_1of[f"{col}{linha_1of}"].value
        nova_aba_modelo.cell(row=linha_dest, column=col_dest, value=valor)

# Função para salvar apenas a aba MODELO
def salvar_apenas_modelo(novo_wb, caminho_arquivo_excel):
    for sheet in novo_wb.sheetnames:
        if sheet != 'MODELO':
            novo_wb.remove(novo_wb[sheet])
    novo_wb.save(caminho_arquivo_excel)

# Função para converter Excel para PDF
def excel_para_pdf(caminho_excel, caminho_pdf):
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    try:
        workbook = excel.Workbooks.Open(os.path.abspath(caminho_excel))
        workbook.ExportAsFixedFormat(0, os.path.abspath(caminho_pdf))
        workbook.Close(False)
    finally:
        excel.Quit()

# Processar as linhas da aba 1OF
total_linhas = aba_1of.max_row
for linha in range(2, total_linhas + 1):
    if all(cell.value is None for cell in aba_1of[linha]):
        continue
    
    novo_wb = load_workbook(CAMINHO_ENTRADA, keep_vba=True)
    nova_aba_modelo = novo_wb['MODELO']
    transferir_dados(linha, nova_aba_modelo)
    
    # Pegando o número da matrícula da coluna B (linha atual)
    matricula = aba_1of.cell(row=linha, column=2).value  # Número da matrícula
    
    # Usando o número da matrícula no nome do arquivo
    nome_arquivo = f"{matricula}"
    caminho_excel = os.path.join(CAMINHO_SAIDA, f"{nome_arquivo}.xlsm")
    caminho_pdf = os.path.join(CAMINHO_SAIDA, f"{nome_arquivo}.pdf")
    
    salvar_apenas_modelo(novo_wb, caminho_excel)
    excel_para_pdf(caminho_excel, caminho_pdf)
    print(f"Arquivos salvos: {caminho_excel} e {caminho_pdf}")

print("Processo concluído! Todos os arquivos foram salvos.")
