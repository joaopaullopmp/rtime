import streamlit as st
from streamlit_calendar import calendar
import pandas as pd
from datetime import datetime
from database_manager import DatabaseManager

def prepare_calendar_events(timesheet_data, users_df, projects_df, absences_df, current_user_info):
    """
    Prepara eventos do calendário com restrições baseadas no papel do usuário
    Inclui horas faturáveis, não faturáveis e feriados
    """
    events = []
    
    # Filtrar dados baseado no papel do usuário
    if current_user_info['role'].lower() != 'admin':
        timesheet_data = timesheet_data[timesheet_data['user_id'] == current_user_info['user_id']]
        # Não filtramos feriados por usuário - eles serão exibidos para todos
    
    # Processar registros de horas
    for _, row in timesheet_data.iterrows():
        try:
            # Buscar nome do usuário
            user_query = users_df[users_df['user_id'] == row['user_id']]
            if user_query.empty:
                user_name = "Usuário Desconhecido"
            else:
                user_name = f"{user_query['First_Name'].iloc[0]} {user_query['Last_Name'].iloc[0]}"
            
            # Buscar nome do projeto
            project_query = projects_df[projects_df['project_id'] == row['project_id']]
            if project_query.empty:
                project_name = "Projeto Desconhecido"
            else:
                project_name = project_query['project_name'].iloc[0]
            
            # Converter billable para booleano para garantir a compatibilidade
            is_billable = False
            if 'billable' in row:
                if isinstance(row['billable'], bool):
                    is_billable = row['billable']
                elif isinstance(row['billable'], (int, float)):
                    is_billable = bool(row['billable'])
                elif isinstance(row['billable'], str):
                    is_billable = row['billable'].lower() in ('true', 't', 'yes', 'y', '1')
            
            event = {
                'id': str(row['id']),
                'title': f"{user_name} - {project_name}",
                'start': pd.to_datetime(row['start_date']).strftime('%Y-%m-%d %H:%M:%S'),
                'end': pd.to_datetime(row['end_date']).strftime('%Y-%m-%d %H:%M:%S'),
                'description': str(row.get('description', '')),
                'backgroundColor': '#1E88E5' if is_billable else '#FF4B4B',
                'extendedProps': {
                    'hours': float(row.get('hours', 0)),
                    'billable': 'Sim' if is_billable else 'Não',
                    'overtime': 'Sim' if row.get('overtime', False) else 'Não',
                    'type': 'work'
                }
            }
            events.append(event)
        except Exception as e:
            st.error(f"Erro ao processar registro de horas: {str(e)}")
            continue
    
    # Separar e processar feriados
    try:
        # Filtrar apenas feriados
        holidays_df = absences_df[
            absences_df['absence_type'].notna() & 
            absences_df['absence_type'].astype(str).str.lower().str.contains('feriado')
        ]
        
        # Agrupar feriados por data para evitar duplicatas
        unique_holidays = {}
        for _, row in holidays_df.iterrows():
            start_date = pd.to_datetime(row['start_date']).strftime('%Y-%m-%d')
            
            # Obter descrição do feriado
            holiday_name = str(row.get('description', '')) if row.get('description') is not None else ""
            if not holiday_name:
                holiday_name = "Feriado"
            
            # Usamos a data como chave para evitar duplicatas
            if start_date not in unique_holidays:
                unique_holidays[start_date] = holiday_name
        
        # Criar eventos para cada feriado único
        for date, name in unique_holidays.items():
            event = {
                'id': f"holiday_{date}",
                'title': name,  # Apenas o nome do feriado
                'start': date,
                'end': date,
                'backgroundColor': '#4CAF50',  # Verde para feriados
                'textColor': '#FFFFFF',
                'allDay': True,
                'display': 'background',  # Exibe como fundo para destacar o dia todo
                'extendedProps': {
                    'type': 'holiday'
                }
            }
            events.append(event)
        
        # Processar outras ausências (não feriados)
        non_holidays = absences_df[
            ~(absences_df['absence_type'].notna() & 
              absences_df['absence_type'].astype(str).str.lower().str.contains('feriado'))
        ]
        
        # Filtrar ausências não-feriados por usuário
        if current_user_info['role'].lower() != 'admin':
            non_holidays = non_holidays[non_holidays['user_id'] == current_user_info['user_id']]
        
        # Processar ausências regulares (não feriados)
        for _, row in non_holidays.iterrows():
            try:
                # Buscar nome do usuário
                user_query = users_df[users_df['user_id'] == row['user_id']]
                if user_query.empty:
                    user_name = "Usuário Desconhecido"
                else:
                    user_name = f"{user_query['First_Name'].iloc[0]} {user_query['Last_Name'].iloc[0]}"
                
                # Obter tipo de ausência com valor padrão para o título
                abs_type = str(row.get('absence_type', '')) if row.get('absence_type') is not None else "Ausência"
                
                event = {
                    'id': f"absence_{row['absence_id']}",
                    'title': f"{user_name} - {abs_type}",
                    'start': pd.to_datetime(row['start_date']).strftime('%Y-%m-%d'),
                    'end': pd.to_datetime(row['end_date']).strftime('%Y-%m-%d'),
                    'description': str(row.get('description', '')),
                    'backgroundColor': '#FFC107',  # Amarelo para ausências
                    'allDay': True,
                    'extendedProps': {
                        'type': 'absence'
                    }
                }
                events.append(event)
            except Exception as e:
                st.error(f"Erro ao processar ausência (ID: {row.get('absence_id', 'desconhecido')}): {str(e)}")
                continue
                
    except Exception as e:
        st.error(f"Erro ao processar feriados/ausências: {str(e)}")
    
    return events

def calendar_page():
    st.title("Calendário de Horas")
    
    try:
        # Carregar dados do banco SQLite em vez de arquivos Excel
        db_manager = DatabaseManager()
        
        # Consultar tabelas necessárias
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        projects_df = db_manager.query_to_df("SELECT * FROM projects")
        absences_df = db_manager.query_to_df("SELECT * FROM absences")
        
        # Verificar o papel do usuário atual
        is_admin = st.session_state.user_info['role'].lower() == 'admin'
        current_user_info = st.session_state.user_info
        
        # Filtros - mostrar todos apenas para admin
        if is_admin:
            col1, col2 = st.columns(2)
            with col1:
                # Simplificar o filtro de usuários
                user_list = []
                for _, user in users_df.iterrows():
                    user_list.append(f"{user['First_Name']} {user['Last_Name']}")
                
                selected_users = st.multiselect(
                    "Filtrar por Usuário",
                    options=user_list
                )
            
            with col2:
                # Simplificar o filtro de projetos
                selected_projects = st.multiselect(
                    "Filtrar por Projeto",
                    options=projects_df['project_name'].unique().tolist()
                )
            
            # Aplicar filtros para admin
            filtered_data = timesheet_df.copy()
            
            if selected_users:
                user_ids = []
                for user_name in selected_users:
                    first_name, last_name = user_name.split(' ', 1)
                    user = users_df[
                        (users_df['First_Name'] == first_name) & 
                        (users_df['Last_Name'] == last_name)
                    ]
                    if not user.empty:
                        user_ids.append(user['user_id'].iloc[0])
                filtered_data = filtered_data[filtered_data['user_id'].isin(user_ids)]
            
            if selected_projects:
                project_ids = []
                for project_name in selected_projects:
                    project = projects_df[projects_df['project_name'] == project_name]
                    if not project.empty:
                        project_ids.append(project['project_id'].iloc[0])
                filtered_data = filtered_data[filtered_data['project_id'].isin(project_ids)]
        else:
            # Para usuários normais, filtrar apenas seus próprios dados
            filtered_data = timesheet_df[timesheet_df['user_id'] == current_user_info['user_id']]
        
        # Converter dados para eventos
        events = prepare_calendar_events(filtered_data, users_df, projects_df, absences_df, current_user_info)
        
        # Configuração do calendário
        calendar_options = {
            "headerToolbar": {
                "left": "today prev,next",
                "center": "title",
                "right": "dayGridMonth,timeGridWeek,timeGridDay"
            },
            "initialView": "timeGridWeek",
            "slotMinTime": "08:00:00",
            "slotMaxTime": "20:00:00",
            "slotDuration": "00:30:00",
            "height": 700,
            "eventTimeFormat": {
                "hour": "2-digit",
                "minute": "2-digit",
                "hour12": False
            }
        }
        
        # Renderizar calendário
        calendar(events=events, options=calendar_options)
        
        # Legenda com três categorias
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("🔵 Horas Faturáveis")
        with col2:
            st.markdown("🔴 Horas Não Faturáveis")
        with col3:
            st.markdown("🟢 Feriados")
            
    except Exception as e:
        st.error(f"Erro ao carregar calendário: {str(e)}")
        import traceback
        st.error(traceback.format_exc())