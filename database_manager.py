import sqlite3
import pandas as pd
import os
from contextlib import contextmanager
from datetime import datetime
import hashlib


class DatabaseManager:
    def __init__(self, db_file='timetracker.db'):
        """Inicializa o gerenciador de banco de dados SQLite"""
        self.db_file = db_file
        # Cria o banco de dados, se não existir
        self._initialize_db()
    
    def _initialize_db(self):
        """Inicializa o banco de dados com as tabelas necessárias"""
        with self._get_connection() as conn:
            # O banco de dados já deve ter sido criado pelo script de migração
            pass
    
    @contextmanager
    def _get_connection(self):
        """Retorna uma conexão com o banco de dados"""
        conn = None
        try:
            conn = sqlite3.connect(self.db_file)
            yield conn
        finally:
            if conn:
                conn.close()
    
    def execute_query(self, query, params=None):
        """Executa uma query SQL com parâmetros opcionais"""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            return cursor
    
    def fetch_all(self, query, params=None):
        """Executa uma query e retorna todos os resultados"""
        cursor = self.execute_query(query, params)
        return cursor.fetchall()
    
    def fetch_one(self, query, params=None):
        """Executa uma query e retorna um único resultado"""
        cursor = self.execute_query(query, params)
        return cursor.fetchone()
    
    def query_to_df(self, query, params=None):
        """Executa uma query e retorna um DataFrame pandas com tipos corrigidos"""
        with self._get_connection() as conn:
            if params:
                df = pd.read_sql_query(query, conn, params=params)
            else:
                df = pd.read_sql_query(query, conn)
            
            # Sanitizar dados e corrigir tipos
            for col in df.columns:
                # Para colunas que normalmente são inteiros
                if col.endswith('_id') or col == 'id':
                    try:
                        # Tenta converter para numérico
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                        # Substitui NaN por 0
                        df[col] = df[col].fillna(0).astype('Int64')  # tipo Int64 permite NaN
                    except Exception as e:
                        print(f"Erro ao converter coluna {col}: {e}")
                
                # Para colunas booleanas
                elif col in ['active', 'billable', 'overtime', 'approved']:
                    try:
                        df[col] = df[col].astype('bool')
                    except:
                        # Se falhar, tenta converter para int primeiro
                        try:
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                            df[col] = df[col].fillna(0).astype('bool')
                        except Exception as e:
                            print(f"Erro ao converter coluna booleana {col}: {e}")
            
            return df
        
class TimesheetManagerSQL:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria um novo registro de timesheet"""
        # Adicionar timestamp atual para created_at e updated_at se não existirem
        from datetime import datetime
        current_time = datetime.now().isoformat()
        
        if 'created_at' not in data:
            data['created_at'] = current_time
        if 'updated_at' not in data:
            data['updated_at'] = current_time
        
        # Preparar as colunas e valores para a query
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO timesheet ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        # Retorna o ID do novo registro
        return cursor.lastrowid
    
    def update(self, id, data):
        """Atualiza um registro existente"""
        # Preparar os pares coluna=valor para a query
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE timesheet SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
        
        # Adicionar o ID no final dos parâmetros
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui um registro"""
        query = "DELETE FROM timesheet WHERE id = ?"
        self.db.execute_query(query, (id,))
        return True
    
    def get_user_entries(self, user_id, start_date=None, end_date=None):
        """Obtém registros de um usuário com filtro de datas opcional"""
        query = "SELECT * FROM timesheet WHERE user_id = ?"
        params = [user_id]
        
        if start_date:
            query += " AND start_date >= ?"
            params.append(start_date)
        
        if end_date:
            query += " AND end_date <= ?"
            params.append(end_date)
        
        return self.db.query_to_df(query, params)
    
    def get_project_entries(self, project_id):
        """Obtém todos os registros de um projeto específico"""
        query = "SELECT * FROM timesheet WHERE project_id = ?"
        return self.db.query_to_df(query, (project_id,))
    
    def calculate_total_hours(self, entries):
        """Calcula o total de horas para um conjunto de registros"""
        if entries.empty:
            return 0
        return entries['hours'].sum()


class UserManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def hash_password(self, password):
        """Gera hash de senha com SHA-256"""
        return hashlib.sha256(str(password).encode()).hexdigest()
    
    def create(self, data):
        """Cria um novo usuário"""
        # Gerar hash da senha se fornecida
        if 'password' in data:
            data['password'] = self.hash_password(data['password'])
        
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO utilizadores ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê informações de usuários"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM utilizadores")
        else:
            return self.db.query_to_df("SELECT * FROM utilizadores WHERE user_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza um usuário existente"""
        # REMOVER O AUTO-HASH - app.py já envia hash correto
        
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE utilizadores SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui um usuário"""
        query = "DELETE FROM utilizadores WHERE user_id = ?"
        self.db.execute_query(query, (id,))
        return True
    
    def login(self, email, password):
        """Verifica credenciais de login"""
        # Hash da senha fornecida
        hashed_password = self.hash_password(password)
        
        # Buscar usuário pelo email
        query = "SELECT * FROM utilizadores WHERE email = ?"
        user_df = self.db.query_to_df(query, (email,))
        
        if user_df.empty:
            return False, "Email não encontrado"
        
        user = user_df.iloc[0]
        
        # Verificar se usuário está ativo
        if not user['active']:
            return False, "Usuário inativo"
        
        # Verificar senha
        if user['password'] == hashed_password:
            return True, user.to_dict()
        
        return False, "Senha incorreta"


class GroupManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria um novo grupo"""
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO groups ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê informações de grupos"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM groups")
        else:
            return self.db.query_to_df("SELECT * FROM groups WHERE id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza um grupo existente"""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE groups SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui um grupo"""
        query = "DELETE FROM groups WHERE id = ?"
        self.db.execute_query(query, (id,))
        return True


class ClientManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria um novo cliente"""
        # Garantir que group_id seja um número
        if 'group_id' in data and not isinstance(data['group_id'], (int, float)):
            try:
                data['group_id'] = int(data['group_id'])
            except (ValueError, TypeError):
                # Valor padrão caso não seja possível converter
                data['group_id'] = 0
        
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO clients ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê informações de clientes"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM clients")
        else:
            return self.db.query_to_df("SELECT * FROM clients WHERE client_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza um cliente existente"""
        # Garantir que group_id seja um número
        if 'group_id' in data and not isinstance(data['group_id'], (int, float)):
            try:
                data['group_id'] = int(data['group_id'])
            except (ValueError, TypeError):
                # Valor padrão caso não seja possível converter
                data['group_id'] = 0
        
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE clients SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE client_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui um cliente"""
        query = "DELETE FROM clients WHERE client_id = ?"
        self.db.execute_query(query, (id,))
        return True
    
    def get_active_clients(self):
        """Retorna apenas clientes ativos"""
        return self.db.query_to_df("SELECT * FROM clients WHERE active = 1")
    
    def get_clients_by_group(self, group_id):
        """Retorna clientes de um grupo específico"""
        return self.db.query_to_df("SELECT * FROM clients WHERE group_id = ?", (group_id,))


class ProjectManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def ensure_numeric(self, value, default=0):
        """Garante que o valor seja numérico"""
        if isinstance(value, (int, float)):
            return value
        try:
            return int(value)
        except (ValueError, TypeError):
            return default
    
    def create(self, data):
        """Cria um novo projeto"""
        # Garantir que IDs sejam numéricos
        if 'client_id' in data:
            data['client_id'] = self.ensure_numeric(data['client_id'])
        if 'group_id' in data:
            data['group_id'] = self.ensure_numeric(data['group_id'])
        
        # Converter valores numéricos para tipos adequados
        if 'hourly_rate' in data:
            try:
                data['hourly_rate'] = float(data['hourly_rate'])
            except (ValueError, TypeError):
                data['hourly_rate'] = 0.0
                
        if 'total_hours' in data:
            try:
                data['total_hours'] = float(data['total_hours'])
            except (ValueError, TypeError):
                data['total_hours'] = 0.0
                
        if 'total_cost' in data:
            try:
                data['total_cost'] = float(data['total_cost'])
            except (ValueError, TypeError):
                data['total_cost'] = 0.0
        
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO projects ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def update(self, id, data):
        """Atualiza um projeto existente"""
        # Garantir que IDs sejam numéricos
        if 'client_id' in data:
            data['client_id'] = self.ensure_numeric(data['client_id'])
        if 'group_id' in data:
            data['group_id'] = self.ensure_numeric(data['group_id'])
        
        # Converter valores numéricos para tipos adequados
        if 'hourly_rate' in data:
            try:
                data['hourly_rate'] = float(data['hourly_rate'])
            except (ValueError, TypeError):
                data['hourly_rate'] = 0.0
        
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE projects SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE project_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def read(self, id=None):
        """Lê informações de projetos"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM projects")
        else:
            return self.db.query_to_df("SELECT * FROM projects WHERE project_id = ?", (id,))
    
    def delete(self, id):
        """Exclui um projeto"""
        query = "DELETE FROM projects WHERE project_id = ?"
        self.db.execute_query(query, (id,))
        return True
    
    def get_active_projects(self):
        """Retorna apenas projetos ativos"""
        return self.db.query_to_df("SELECT * FROM projects WHERE status IN ('active', 'Active')")
    
    def get_projects_by_client(self, client_id):
        """Retorna projetos de um cliente específico"""
        return self.db.query_to_df("SELECT * FROM projects WHERE client_id = ?", (client_id,))
    
    def get_active_projects_by_client(self, client_id):
        """Retorna projetos ativos de um cliente específico"""
        return self.db.query_to_df(
            "SELECT * FROM projects WHERE client_id = ? AND status IN ('active', 'Active')",
            (client_id,)
        )


class TaskCategoryManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria uma nova categoria de tarefa"""
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO task_categories ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê informações de categorias de tarefa"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM task_categories")
        else:
            return self.db.query_to_df("SELECT * FROM task_categories WHERE task_category_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza uma categoria de tarefa existente"""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE task_categories SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE task_category_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui uma categoria de tarefa"""
        query = "DELETE FROM task_categories WHERE task_category_id = ?"
        self.db.execute_query(query, (id,))
        return True


class ActivityManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria uma nova atividade"""
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO activities ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê informações de atividades"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM activities")
        else:
            return self.db.query_to_df("SELECT * FROM activities WHERE activity_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza uma atividade existente"""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE activities SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE activity_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui uma atividade"""
        query = "DELETE FROM activities WHERE activity_id = ?"
        self.db.execute_query(query, (id,))
        return True


class RateManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria uma nova rate"""
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO rates ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê informações de rates"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM rates")
        else:
            return self.db.query_to_df("SELECT * FROM rates WHERE rate_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza uma rate existente"""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE rates SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE rate_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui uma rate"""
        query = "DELETE FROM rates WHERE rate_id = ?"
        self.db.execute_query(query, (id,))
        return True


class AbsenceManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria um novo registro de ausência"""

         # Adicionar timestamp atual para created_at e updated_at se não existirem
        from datetime import datetime
        current_time = datetime.now().isoformat()
        
        if 'created_at' not in data:
            data['created_at'] = current_time
        if 'updated_at' not in data:
            data['updated_at'] = current_time
            
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO absences ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê registros de ausência"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM absences")
        else:
            return self.db.query_to_df("SELECT * FROM absences WHERE absence_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza um registro de ausência"""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE absences SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE absence_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui um registro de ausência"""
        query = "DELETE FROM absences WHERE absence_id = ?"
        self.db.execute_query(query, (id,))
        return True
    
    def get_user_absences(self, user_id):
        """Obtém ausências de um usuário"""
        query = "SELECT * FROM absences WHERE user_id = ?"
        return self.db.query_to_df(query, (user_id,))


class TravelExpenseManager:
    def __init__(self):
        self.db = DatabaseManager()
    
    def create(self, data):
        """Cria um novo registro de despesa de viagem"""

        # Adicionar timestamp atual para created_at e updated_at se não existirem
        from datetime import datetime
        current_time = datetime.now().isoformat()
        
        if 'created_at' not in data:
            data['created_at'] = current_time
        if 'updated_at' not in data:
            data['updated_at'] = current_time

        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO travel_expenses ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def read(self, id=None):
        """Lê registros de despesa de viagem"""
        if id is None:
            return self.db.query_to_df("SELECT * FROM travel_expenses")
        else:
            return self.db.query_to_df("SELECT * FROM travel_expenses WHERE travel_id = ?", (id,))
    
    def update(self, id, data):
        """Atualiza um registro de despesa de viagem"""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE travel_expenses SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE travel_id = ?"
        
        params = list(data.values())
        params.append(id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete(self, id):
        """Exclui um registro de despesa de viagem"""
        query = "DELETE FROM travel_expenses WHERE travel_id = ?"
        self.db.execute_query(query, (id,))
        return True
    
    def get_user_expenses(self, user_id):
        """Obtém despesas de viagem de um usuário"""
        query = "SELECT * FROM travel_expenses WHERE user_id = ?"
        return self.db.query_to_df(query, (user_id,))