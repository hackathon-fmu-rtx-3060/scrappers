import pandas as pd
from tabula import read_pdf

# Caminho do arquivo PDF
pdf_file = "cofap.pdf"

# Páginas para extrair a tabela
paginas = "86"

# Extraindo tabelas das páginas especificadas
try:
    # Extraímos todas as tabelas encontradas nas páginas definidas
    tabelas = read_pdf(pdf_file, pages=paginas, multiple_tables=True, lattice=True,stream=False)

    # Processar cada tabela extraída
    tabela_processada = []
    for tabela in tabelas:
        # Preencher células mescladas
        tabela = tabela.ffill(axis=0)  # Preenche células mescladas verticalmente
        tabela_processada.append(tabela)

    # Concatenar as tabelas processadas
    tabela_final = pd.concat(tabela_processada, ignore_index=False)

    # Salvar como Excel
    tabela_final.to_excel("tabela_processada.xlsx", index=False)

    print("Tabela extraída e salva em 'tabela_processada.xlsx'.")
except Exception as e:
    print(f"Erro ao processar o arquivo: {e}")
