import pandas as pd
import sqlite3
import os

# Caminho para o banco de dados - ajuste conforme necessário
DB_PATH = 'timetracker.db'
table_name = 'utilizadores'

def read_table_from_db(table_name):
    """
    Lê uma tabela do banco de dados SQLite e retorna como DataFrame.
    Substitui a funcionalidade pd.read_excel()
    """
    conn = sqlite3.connect(DB_PATH)
    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def save_table_to_db(df, table_name):
    """
    Salva um DataFrame em uma tabela do banco de dados SQLite.
    Substitui a funcionalidade df.to_excel()
    """
    conn = sqlite3.connect(DB_PATH)
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    conn.close()
    return True

def execute_query(query, params=None):
    """
    Executa uma query SQL personalizada
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    if params:
        cursor.execute(query, params)
    else:
        cursor.execute(query)
    
    conn.commit()
    result = cursor.fetchall()
    conn.close()
    return result

def get_table_columns(table_name):
    """
    Retorna a lista de colunas de uma tabela
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = [info[1] for info in cursor.fetchall()]
    conn.close()
    return columns