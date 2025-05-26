# project_phase_manager.py
import streamlit as st
import pandas as pd
from datetime import datetime
from database_manager import DatabaseManager

class ProjectPhaseManager:
    def __init__(self):
        self.db = DatabaseManager()
        self._ensure_table_exists()
    
    def _ensure_table_exists(self):
        """Garante que a tabela project_phases existe no banco de dados"""
        query = """
        CREATE TABLE IF NOT EXISTS project_phases (
            phase_id INTEGER PRIMARY KEY,
            project_id INTEGER NOT NULL,
            phase_name TEXT NOT NULL,
            phase_description TEXT,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            total_hours REAL NOT NULL,
            total_cost REAL NOT NULL,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (project_id) REFERENCES projects(project_id)
        );
        """
        self.db.execute_query(query)
    
    def create(self, data):
        """Cria uma nova fase de projeto"""
        # Verificar se os valores não excedem os limites do projeto
        if not self._validate_phase_values(data):
            return False, "Os valores da fase excedem os limites disponíveis do projeto."
        
        # Adicionar timestamp atual para created_at e updated_at se não existirem
        current_time = datetime.now().isoformat()
        if 'created_at' not in data:
            data['created_at'] = current_time
        if 'updated_at' not in data:
            data['updated_at'] = current_time
        
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO project_phases ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return True, cursor.lastrowid
    
    def read(self, phase_id=None, project_id=None):
        """Lê fases de projeto"""
        if phase_id is not None:
            return self.db.query_to_df("SELECT * FROM project_phases WHERE phase_id = ?", (phase_id,))
        elif project_id is not None:
            return self.db.query_to_df("SELECT * FROM project_phases WHERE project_id = ? ORDER BY start_date", (project_id,))
        else:
            return self.db.query_to_df("SELECT * FROM project_phases ORDER BY project_id, start_date")
    
    def update(self, phase_id, data):
        """Atualiza uma fase de projeto existente"""
        # Verificar se os valores não excedem os limites do projeto
        if not self._validate_phase_values(data, phase_id):
            return False, "Os valores da fase excedem os limites disponíveis do projeto."
        
        # Adicionar timestamp atual para updated_at
        data['updated_at'] = datetime.now().isoformat()
        
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE project_phases SET {set_clause} WHERE phase_id = ?"
        
        params = list(data.values())
        params.append(phase_id)
        
        self.db.execute_query(query, params)
        return True, "Fase atualizada com sucesso."
    
    def delete(self, phase_id):
        """Exclui uma fase de projeto"""
        query = "DELETE FROM project_phases WHERE phase_id = ?"
        self.db.execute_query(query, (phase_id,))
        return True

    def get_project_phases_summary(self, project_id):
        """Retorna um resumo das fases de um projeto"""
        phases = self.read(project_id=project_id)
        if phases.empty:
            return {
                'total_phases': 0,
                'total_hours': 0,
                'total_cost': 0
            }
        
        return {
            'total_phases': len(phases),
            'total_hours': phases['total_hours'].sum(),
            'total_cost': phases['total_cost'].sum()
        }
    
    def _validate_phase_values(self, data, phase_id=None):
        """
        Valida se os valores da fase não excedem os limites do projeto
        Considera outras fases existentes exceto a própria fase sendo editada
        """
        project_id = data.get('project_id')
        if not project_id:
            # Se estamos atualizando, precisamos pegar o project_id da fase existente
            if phase_id:
                phase_data = self.read(phase_id=phase_id)
                if not phase_data.empty:
                    project_id = phase_data['project_id'].iloc[0]
                else:
                    return False
            else:
                return False
        
        # Buscar dados do projeto
        project_query = "SELECT total_hours, total_cost FROM projects WHERE project_id = ?"
        project_data = self.db.query_to_df(project_query, (project_id,))
        
        if project_data.empty:
            return False
        
        project_total_hours = float(project_data['total_hours'].iloc[0])
        project_total_cost = float(project_data['total_cost'].iloc[0])
        
        # Buscar todas as fases do projeto exceto a fase atual (se estiver editando)
        if phase_id:
            phases_query = "SELECT total_hours, total_cost FROM project_phases WHERE project_id = ? AND phase_id != ?"
            phases_data = self.db.query_to_df(phases_query, (project_id, phase_id))
        else:
            phases_query = "SELECT total_hours, total_cost FROM project_phases WHERE project_id = ?"
            phases_data = self.db.query_to_df(phases_query, (project_id,))
        
        # Calcular o total de horas e custo já alocados em outras fases
        existing_total_hours = phases_data['total_hours'].sum() if not phases_data.empty else 0
        existing_total_cost = phases_data['total_cost'].sum() if not phases_data.empty else 0
        
        # Verificar se ao adicionar/atualizar esta fase, os totais excedem os limites do projeto
        new_phase_hours = float(data.get('total_hours', 0))
        new_phase_cost = float(data.get('total_cost', 0))
        
        # Verificar se os novos valores excedem os limites do projeto
        if existing_total_hours + new_phase_hours > project_total_hours:
            return False
        
        if existing_total_cost + new_phase_cost > project_total_cost:
            return False
        
        return True
        
    def get_available_resources(self, project_id, phase_id=None):
        """
        Calcula os recursos disponíveis (horas e custo) para alocação em uma nova fase
        ou para edição de uma fase existente
        """
        # Buscar dados do projeto
        project_query = "SELECT total_hours, total_cost FROM projects WHERE project_id = ?"
        project_data = self.db.query_to_df(project_query, (project_id,))
        
        if project_data.empty:
            return None
        
        project_total_hours = float(project_data['total_hours'].iloc[0])
        project_total_cost = float(project_data['total_cost'].iloc[0])
        
        # Buscar todas as fases do projeto exceto a fase atual (se estiver editando)
        if phase_id:
            phases_query = "SELECT total_hours, total_cost FROM project_phases WHERE project_id = ? AND phase_id != ?"
            phases_data = self.db.query_to_df(phases_query, (project_id, phase_id))
        else:
            phases_query = "SELECT total_hours, total_cost FROM project_phases WHERE project_id = ?"
            phases_data = self.db.query_to_df(phases_query, (project_id,))
        
        # Calcular o total de horas e custo já alocados em outras fases
        allocated_hours = phases_data['total_hours'].sum() if not phases_data.empty else 0
        allocated_cost = phases_data['total_cost'].sum() if not phases_data.empty else 0
        
        # Calcular recursos disponíveis
        available_hours = project_total_hours - allocated_hours
        available_cost = project_total_cost - allocated_cost
        
        return {
            'project_total_hours': project_total_hours,
            'project_total_cost': project_total_cost,
            'allocated_hours': allocated_hours,
            'allocated_cost': allocated_cost,
            'available_hours': available_hours,
            'available_cost': available_cost
        }