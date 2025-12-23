import sqlite3
import pandas as pd
from pathlib import Path

def list_files(folder_path="."):
    folder = Path(folder_path)
    files = [
        f.name for f in folder.iterdir()
        if f.is_file() and not f.name.startswith('~') and not f.name.endswith('.tmp')
    ]
    return files

def drop_table_if_exists(table_name, conn):
    cursor = conn.cursor()
    cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
    conn.commit()
    print(f"Table '{table_name}' dropped.")


excels = list_files('outputs')

conn = sqlite3.connect('db.sqlite3')


drop_table_if_exists('raw_ssf',conn)

for x in excels:
	ruta = f"outputs/{x}"
	df = pd.read_excel(ruta, sheet_name=0)
	df.to_sql('raw_ssf', conn, if_exists='append', index=False)
	print(f"The data in {x} was inserted.")

conn.close()






