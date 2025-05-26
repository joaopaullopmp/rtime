import sqlite3
import os

def clean_database(db_file='timetracker.db'):
    """Limpa as tabelas do banco de dados para nova migração"""
    if os.path.exists(db_file):
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()

        # Lista todas as tabelas que podem ter sido parcialmente migradas
        tables = ["timesheet", "absences", "travel_expenses"]
        
        for table in tables:
            try:
                # Remover a tabela se existir
                cursor.execute(f"DROP TABLE IF EXISTS {table}")
                print(f"Tabela {table} removida com sucesso.")
            except Exception as e:
                print(f"Erro ao remover tabela {table}: {e}")
        
        conn.commit()
        conn.close()
        print("Limpeza do banco de dados concluída.")
    else:
        print(f"Banco de dados {db_file} não encontrado.")

if __name__ == "__main__":
    clean_database()