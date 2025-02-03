import os
from openpyxl import load_workbook
from datetime import datetime

# Caminhos dos arquivos
CAMINHO_ENTRADA = os.path.join('..', 'dados', 'entrada', 'exemplo.xlsm')  # Arquivo de entrada
CAMINHO_SAIDA = os.path.join('..', 'dados', 'saida')                     # Pasta de saída
PASTA_LOGS = os.path.join('..', 'logs', 'execucao.log')                  # Arquivo de log

# Criar pasta de saída se não existir
os.makedirs(CAMINHO_SAIDA, exist_ok=True)

# Função para registrar logs
def registrar_log(mensagem):
    with open(PASTA_LOGS, 'a') as log:
        log.write(f"{datetime.now()} - {mensagem}\n")

# Carregar o arquivo Excel com openpyxl
try:
    wb = load_workbook(CAMINHO_ENTRADA)
    aba_1of = wb['1OF']
    aba_modelo = wb['MODELO']
    registrar_log("Arquivo exemplo.xlsm carregado com sucesso.")
except Exception as e:
    registrar_log(f"Erro ao carregar o arquivo exemplo.xlsm: {e}")
    exit()

# Função para atualizar as fórmulas na aba MODELO
def atualizar_formulas(aba_modelo, linha):
    for row in aba_modelo.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):  # Verificar se é uma fórmula
                formula = cell.value
                if "'1OF'!" in formula:  # Verificar se a fórmula referencia a aba 1OF
                    # Atualizar a referência da linha na fórmula
                    nova_formula = formula.replace(f"'{linha-1}'", f"'{linha}'")
                    cell.value = nova_formula

# Processar cada linha da aba "1OF"
for indice, linha in enumerate(aba_1of.iter_rows(min_row=2, values_only=True), start=2):
    try:
        # Criar um novo arquivo Excel
        novo_wb = load_workbook(CAMINHO_ENTRADA)
        nova_aba_modelo = novo_wb['MODELO']

        # Atualizar as fórmulas na aba MODELO para referenciar a linha correta
        atualizar_formulas(nova_aba_modelo, indice)

        # Salvar apenas a aba MODELO no novo arquivo
        novo_wb.remove(novo_wb['1OF'])  # Remover a aba 1OF
        novo_wb.save(os.path.join(CAMINHO_SAIDA, f'MODELO_linha_{indice-1}.xlsx'))

        registrar_log(f"Arquivo salvo: MODELO_linha_{indice-1}.xlsx")
    except Exception as e:
        registrar_log(f"Erro ao processar linha {indice-1}: {e}")

registrar_log("Processo concluído! Todos os arquivos foram salvos.")