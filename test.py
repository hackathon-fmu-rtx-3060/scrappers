import pandas as pd

# Caminho do arquivo Excel
file_path = 'tabela_processada.xlsx'

# Verificar as abas disponíveis
sheet_names = pd.ExcelFile(file_path, engine='openpyxl').sheet_names
print("Abas disponíveis:", sheet_names)
