import pandas as pd
from sqlalchemy import create_engine

# Caminho do arquivo Excel
file_path = 'tabela_processada.xlsx'

# URL de conexão ao banco de dados
db_url = "postgresql://neondb_owner:vRILiN7HOA5M@ep-red-band-a5vcpjq2.us-east-2.aws.neon.tech/neondb?sslmode=require"

# Nome da tabela que será populada com a aba "backup"
backup_table_name = "dados_backup"

# Ler a aba "backup" do arquivo Excel (especificando o engine)
backup_data = pd.read_excel(file_path, sheet_name='BACKUP', engine='openpyxl')

# Criar conexão com o banco de dados
engine = create_engine(db_url)

# Escrever os dados da aba "backup" no banco de dados
backup_data.to_sql(backup_table_name, engine, if_exists='replace', index=False)

print(f"Dados da aba 'backup' inseridos na tabela '{backup_table_name}' com sucesso!")
