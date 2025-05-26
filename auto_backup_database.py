import os
import shutil
import datetime
import time
import logging
import sqlite3
import schedule

# Configurar logging
logging.basicConfig(
    filename='backup_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def backup_database():
    """
    Realiza o backup do banco de dados SQLite para o diretório especificado
    """
    try:
        # Definir caminhos
        source_file = 'timetracker.db'
        backup_dir = r'c:\BKP\DB'
        
        # Verificar se o arquivo de origem existe
        if not os.path.exists(source_file):
            logging.error(f"Arquivo de origem {source_file} não encontrado.")
            return False
        
        # Criar diretório de backup se não existir
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            logging.info(f"Diretório de backup {backup_dir} criado.")
        
        # Criar nome do arquivo de backup com timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"timetracker_backup_{timestamp}.db"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        # Verificar se o banco de dados não está bloqueado
        try:
            conn = sqlite3.connect(source_file)
            conn.close()
        except sqlite3.Error as e:
            logging.error(f"Não foi possível conectar ao banco de dados: {e}")
            return False
        
        # Copiar o arquivo
        shutil.copy2(source_file, backup_path)
        logging.info(f"Backup realizado com sucesso para {backup_path}")
        
        # Manter apenas os 10 backups mais recentes
        cleanup_old_backups(backup_dir, 10)
        
        return True
    
    except Exception as e:
        logging.error(f"Erro durante o backup: {e}")
        return False

def cleanup_old_backups(backup_dir, keep_count=10):
    """
    Remove backups antigos, mantendo apenas os 'keep_count' mais recentes
    """
    try:
        # Listar todos os arquivos de backup
        backup_files = [f for f in os.listdir(backup_dir) 
                        if f.startswith("timetracker_backup_") and f.endswith(".db")]
        
        # Ordenar por data (mais recentes primeiro)
        backup_files.sort(reverse=True)
        
        # Remover arquivos antigos
        for old_file in backup_files[keep_count:]:
            os.remove(os.path.join(backup_dir, old_file))
            logging.info(f"Backup antigo removido: {old_file}")
    
    except Exception as e:
        logging.error(f"Erro ao limpar backups antigos: {e}")

# Agendar backup a cada hora
schedule.every(1).hour.do(backup_database)

# Executar um backup imediatamente na inicialização
backup_database()
print("Serviço de backup iniciado. Pressione Ctrl+C para interromper.")

# Loop principal
while True:
    schedule.run_pending()
    time.sleep(60)  # Verifica a cada minuto