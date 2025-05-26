import os
import sys
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import time
import logging
import datetime
import shutil
import sqlite3
import schedule

class BackupService(win32serviceutil.ServiceFramework):
    _svc_name_ = "TimeTrackerBackup"
    _svc_display_name_ = "TimeTracker Database Backup Service"
    _svc_description_ = "Realiza backup automático do banco de dados TimeTracker a cada hora"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        
        # Configurar logging para o serviço
        log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backup_service.log')
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def SvcStop(self):
        logging.info('Recebendo sinal de parada.')
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        logging.info('Iniciando serviço de backup.')
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        self.main()

    def main(self):
        # Realizar primeiro backup imediatamente
        self.backup_database()
        
        # Configurar schedule para backup a cada hora
        schedule.every(1).hour.do(self.backup_database)
        
        # Loop principal do serviço
        while self.is_alive:
            schedule.run_pending()
            # Verificar a cada 30 segundos se há tarefas pendentes ou sinal de parada
            if win32event.WaitForSingleObject(self.hWaitStop, 30000) == win32event.WAIT_OBJECT_0:
                break
            logging.info("Serviço em execução, aguardando próximo backup agendado.")
    
    def backup_database(self):
        """
        Realiza o backup do banco de dados SQLite para o diretório especificado
        """
        try:
            logging.info("Iniciando processo de backup...")
            
            # Definir caminhos
            script_dir = os.path.dirname(os.path.abspath(__file__))
            source_file = os.path.join(script_dir, 'timetracker.db')
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
            self.cleanup_old_backups(backup_dir, 840)
            
            return True
        
        except Exception as e:
            logging.error(f"Erro durante o backup: {e}")
            return False

    def cleanup_old_backups(self, backup_dir, keep_count=840):
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

if __name__ == '__main__':
    if len(sys.argv) == 1:
        try:
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(BackupService)
            servicemanager.StartServiceCtrlDispatcher()
        except Exception as e:
            print(f"Erro ao iniciar o dispatcher: {e}")
    else:
        win32serviceutil.HandleCommandLine(BackupService)