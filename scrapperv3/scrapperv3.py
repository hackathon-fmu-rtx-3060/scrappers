import camelot
import pandas as pd

def extract_tables_from_pdf(pdf_file, output_excel):
    print('hi')
    try:
        # Extraindo tabelas do PDF
        tables = camelot.read_pdf(pdf_file, pages='86,87,88,89,90,91,92')
        # pages='all';pages='1,2,3'
        if tables.n == 0:
            print("Nenhuma tabela encontrada no PDF.")
            return
        
        print(f"Total de tabelas encontradas: {tables.n}")
        
        # Salvando cada tabela em uma planilha do Excel
        writer = pd.ExcelWriter(output_excel, engine='openpyxl')
        for i, table in enumerate(tables):
            df = table.df
            sheet_name = f'Tabela_{i+1}'
            print(f'fazendo tabela {i+1}')
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Tabela {i+1} salva como '{sheet_name}'")

        writer.close()
        print(f"Tabelas extraídas e salvas no arquivo {output_excel}")

    except Exception as e:
        print(f"Erro ao processar o PDF: {e}")

# Configurações
pdf_file = "../cofap.pdf"  # Substitua pelo caminho do seu arquivo PDF
output_excel = "tabelas_extraidas.xlsx"  # Substitua pelo nome do arquivo Excel de saída

# Extraindo tabelas e salvando em Excel
extract_tables_from_pdf(pdf_file, output_excel)
