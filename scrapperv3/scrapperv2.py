from tabula import read_pdf
import pandas as pd

# Nome do arquivo PDF
pdf_file = "cofap.pdf"

# Nome do arquivo de saída
output_file = "tabelas_extraidas.xlsx"

try:
    # Inicializa um escritor Excel para múltiplas planilhas
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        print(f"Tentando extrair tabelas do arquivo '{pdf_file}'...")

        # Loop por todas as páginas do PDF
        page_number = 86
        while True:
            try:
                # Extrai tabelas da página
                tables = read_pdf(pdf_file, pages=page_number, multiple_tables=True)
                if not tables:
                    print(f"Nenhuma tabela encontrada na página {page_number}.")
                    break

                # Salva cada tabela em uma planilha separada
                for i, table in enumerate(tables, start=1):
                    sheet_name = f"Página{page_number}_Tabela{i}"
                    table.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Tabela extraída da página {page_number}, tabela {i}, salva como '{sheet_name}'.")
                page_number += 1
            except ValueError:
                # Interrompe quando não há mais páginas
                print(f"Fim do arquivo PDF alcançado na página {page_number}.")
                break

    print(f"Extração completa! As tabelas foram salvas no arquivo '{output_file}'.")
except Exception as e:
    print(f"Ocorreu um erro ao processar o arquivo PDF: {e}")
