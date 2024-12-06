from tabula import read_pdf
import pandas as pd

# Nome do arquivo PDF
pdf_file = "cofap.pdf"

# Página a ser processada
page_to_extract = 86

# Nome do arquivo de saída
output_file = "tabela_pagina_86.xlsx"

try:
    # Extrai tabelas da página usando Tabula
    print(f"Tentando extrair tabelas da página {page_to_extract} do arquivo '{pdf_file}'...")
    tables = read_pdf(pdf_file, pages=page_to_extract, multiple_tables=True)

    if tables:
        # Processa e salva cada tabela
        for i, table in enumerate(tables, start=1):
            output = f"tabela_{page_to_extract}_{i}.xlsx"
            table.to_excel(output, index=False)
            print(f"Tabela {i} salva com sucesso no arquivo '{output}'.")
    else:
        print(f"Nenhuma tabela encontrada na página {page_to_extract}.")
except Exception as e:
    print(f"Ocorreu um erro ao processar o arquivo PDF: {e}")
