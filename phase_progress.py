# phase_progress.py
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from project_phase_manager import ProjectPhaseManager
from database_manager import DatabaseManager

def phase_progress_dashboard():
    """Dashboard para monitoramento de progresso das fases de projetos"""
    st.title("Monitoramento de Progresso das Fases")
    
    # Inicializar gerenciadores
    phase_manager = ProjectPhaseManager()
    db_manager = DatabaseManager()
    
    # Carregar dados
    projects_df = db_manager.query_to_df("SELECT * FROM projects WHERE status = 'active'")
    clients_df = db_manager.query_to_df("SELECT * FROM clients")
    timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
    
    # Juntar projetos com clientes
    projects_with_clients = projects_df.merge(
        clients_df[['client_id', 'name']],
        on='client_id',
        how='left'
    )
    
    # Criar lista de projetos
    project_options = []
    for _, project in projects_with_clients.iterrows():
        project_options.append({
            'id': project['project_id'],
            'display': f"{project['project_name']} - {project['name']}"
        })
    
    # Seleção de projeto
    selected_project = st.selectbox(
        "Selecione o Projeto",
        options=project_options,
        format_func=lambda x: x['display']
    )
    
    if selected_project:
        project_id = selected_project['id']
        project_info = projects_with_clients[projects_with_clients['project_id'] == project_id].iloc[0]
        
        # Mostrar informações do projeto
        st.subheader(f"Projeto: {project_info['project_name']}")
        st.markdown(f"**Cliente:** {project_info['name']}")
        st.markdown(f"**Período:** {pd.to_datetime(project_info['start_date']).strftime('%d/%m/%Y')} a {pd.to_datetime(project_info['end_date']).strftime('%d/%m/%Y')}")
        
        # Obter fases do projeto
        phases = phase_manager.read(project_id=project_id)
        
        if phases.empty:
            st.info("Este projeto não possui fases definidas. Adicione fases através do menu 'Relatórios > Fases de Projetos'.")
            return
        
        # Obter registros de timesheet para o projeto
        project_timesheet = timesheet_df[timesheet_df['project_id'] == project_id].copy()
        project_timesheet['start_date'] = pd.to_datetime(project_timesheet['start_date'])
        
        # Agrupar horas por data
        daily_hours = project_timesheet.groupby(
            project_timesheet['start_date'].dt.date
        )['hours'].sum().reset_index()
        
        # Data atual
        today = datetime.now().date()
        
        # Calcular progresso de cada fase
        phases_progress = []
        
        for _, phase in phases.iterrows():
            phase_id = phase['phase_id']
            phase_name = phase['phase_name']
            phase_start = pd.to_datetime(phase['start_date']).date()
            phase_end = pd.to_datetime(phase['end_date']).date()
            phase_status = phase['status']
            
            # Horas planejadas para a fase
            phase_planned_hours = float(phase['total_hours'])
            
            # Filtrar horas registradas para o período da fase
            phase_hours = daily_hours[
                (daily_hours['start_date'] >= phase_start) &
                (daily_hours['start_date'] <= phase_end)
            ]['hours'].sum()
            
            # Duração da fase em dias
            phase_duration = (phase_end - phase_start).days + 1
            
            # Calcular progresso esperado com base na data atual
            if today < phase_start:
                expected_progress = 0  # Fase não iniciada
                time_status = "Não Iniciada"
            elif today > phase_end:
                expected_progress = 100  # Prazo encerrado
                time_status = "Prazo Encerrado"
            else:
                days_elapsed = (today - phase_start).days + 1
                expected_progress = min(100, (days_elapsed / phase_duration * 100))
                time_status = "Em Andamento"
            
            # Calcular progresso real com base nas horas
            if phase_planned_hours > 0:
                actual_progress = min(100, (phase_hours / phase_planned_hours * 100))
            else:
                actual_progress = 0
            
            # Determinar estado de conclusão
            if phase_status == 'completed':
                completion_status = "Concluída"
                actual_progress = 100
            elif actual_progress >= 100:
                completion_status = "Horas Excedidas"
            elif actual_progress >= expected_progress:
                completion_status = "No Prazo"
            else:
                completion_status = "Atrasada"
            
            # Calcular diferença entre progresso real e esperado
            progress_diff = actual_progress - expected_progress
            
            # Determinar cor com base no status
            if completion_status == "Concluída":
                status_color = "#4CAF50"  # Verde
            elif completion_status == "No Prazo":
                status_color = "#2196F3"  # Azul
            elif completion_status == "Atrasada":
                status_color = "#FF5722"  # Laranja
            elif completion_status == "Horas Excedidas":
                status_color = "#FFC107"  # Amarelo
            else:
                status_color = "#9E9E9E"  # Cinza
            
            # Adicionar à lista de progresso
            phases_progress.append({
                'phase_id': phase_id,
                'phase_name': phase_name,
                'start_date': phase_start,
                'end_date': phase_end,
                'status': phase_status,
                'planned_hours': phase_planned_hours,
                'actual_hours': phase_hours,
                'expected_progress': expected_progress,
                'actual_progress': actual_progress,
                'progress_diff': progress_diff,
                'time_status': time_status,
                'completion_status': completion_status,
                'status_color': status_color
            })
        
        # Converter para DataFrame
        progress_df = pd.DataFrame(phases_progress)
        
        # Exibir progresso das fases em um visual mais rico
        st.subheader("Progresso das Fases")
        
        # Mostrar cards de status para cada fase
        for i, row in enumerate(progress_df.itertuples()):
            col1, col2 = st.columns([3, 1])
            
            with col1:
                # Card com título da fase e barra de progresso
                st.markdown(f"""
                <div style="padding: 1rem; border-radius: 0.5rem; border-left: 0.5rem solid {row.status_color}; margin-bottom: 1rem;">
                    <h3 style="margin-top: 0; margin-bottom: 0.5rem;">{row.phase_name}</h3>
                    <p style="margin-bottom: 0.5rem;"><span style="color: #666;">Período:</span> {row.start_date.strftime('%d/%m/%Y')} a {row.end_date.strftime('%d/%m/%Y')}</p>
                    <div style="display: flex; align-items: center;">
                        <div style="background-color: #e0e0e0; border-radius: 0.25rem; height: 0.75rem; width: 100%; margin-right: 1rem;">
                            <div style="background-color: {row.status_color}; border-radius: 0.25rem; height: 0.75rem; width: {row.actual_progress}%;"></div>
                        </div>
                        <span style="font-weight: bold; white-space: nowrap;">{row.actual_progress:.1f}%</span>
                    </div>
                    <p style="margin-top: 0.5rem; margin-bottom: 0;"><span style="color: #666;">Status:</span> <span style="font-weight: bold; color: {row.status_color};">{row.completion_status}</span></p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                # Métricas de horas
                st.markdown(f"""
                <div style="padding: 1rem; border-radius: 0.5rem; background-color: #f5f5f5; height: 100%; display: flex; flex-direction: column; justify-content: center;">
                    <p style="margin: 0; text-align: center; font-size: 0.8rem; color: #666;">Horas Planejadas</p>
                    <p style="margin: 0; text-align: center; font-weight: bold; font-size: 1.2rem;">{row.planned_hours:.1f}h</p>
                    <hr style="margin: 0.5rem 0; border: none; border-top: 1px solid #ddd;">
                    <p style="margin: 0; text-align: center; font-size: 0.8rem; color: #666;">Horas Realizadas</p>
                    <p style="margin: 0; text-align: center; font-weight: bold; font-size: 1.2rem;">{row.actual_hours:.1f}h</p>
                </div>
                """, unsafe_allow_html=True)
        
        # Criar gráfico comparativo de progresso
        st.subheader("Comparação de Progresso das Fases")
        
        # Preparar dados para o gráfico
        fig = go.Figure()
        
        # Adicionar barras de progresso esperado
        fig.add_trace(go.Bar(
            y=progress_df['phase_name'],
            x=progress_df['expected_progress'],
            name='Progresso Esperado',
            orientation='h',
            marker=dict(color='rgba(55, 83, 109, 0.7)'),
            hovertemplate='%{x:.1f}%<extra>Esperado</extra>'
        ))
        
        # Adicionar barras de progresso atual com cores diferentes por status
        fig.add_trace(go.Bar(
            y=progress_df['phase_name'],
            x=progress_df['actual_progress'],
            name='Progresso Atual',
            orientation='h',
            marker=dict(
                color=progress_df['status_color']
            ),
            hovertemplate='%{x:.1f}%<extra>Atual</extra>'
        ))
        
        # Ajustar layout
        fig.update_layout(
            title='Comparação entre Progresso Esperado e Atual',
            xaxis_title='Progresso (%)',
            yaxis_title='Fase',
            barmode='group',
            height=400,
            xaxis=dict(range=[0, 110])
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Mostrar tabela detalhada de progresso
        st.subheader("Detalhamento de Progresso das Fases")
        
        # Preparar dados para a tabela
        table_data = progress_df[['phase_name', 'expected_progress', 'actual_progress', 'progress_diff', 'planned_hours', 'actual_hours', 'completion_status']].copy()
        
        # Formatar colunas
        table_data = table_data.rename(columns={
            'phase_name': 'Fase',
            'expected_progress': 'Progresso Esperado (%)',
            'actual_progress': 'Progresso Atual (%)',
            'progress_diff': 'Diferença (%)',
            'planned_hours': 'Horas Planejadas',
            'actual_hours': 'Horas Realizadas',
            'completion_status': 'Status'
        })
        
        # Exibir tabela
        st.dataframe(table_data.set_index('Fase'), use_container_width=True)
        
        # Resumo geral do projeto
        st.subheader("Resumo do Projeto")
        
        # Calcular métricas gerais
        total_phases = len(progress_df)
        completed_phases = sum(progress_df['status'] == 'completed')
        on_track_phases = sum((progress_df['status'] != 'completed') & (progress_df['actual_progress'] >= progress_df['expected_progress']))
        delayed_phases = sum((progress_df['status'] != 'completed') & (progress_df['actual_progress'] < progress_df['expected_progress']))
        overall_progress = progress_df['actual_progress'].mean()
        
        # Mostrar métricas em colunas
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Conclusão Geral", f"{overall_progress:.1f}%")
        
        with col2:
            st.metric("Fases Concluídas", f"{completed_phases}/{total_phases}")
        
        with col3:
            st.metric("Fases no Prazo", f"{on_track_phases}/{total_phases - completed_phases}" if total_phases > completed_phases else "N/A")
        
        # Gráfico de distribuição por status
        status_counts = progress_df['completion_status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Quantidade']
        
        fig_status = px.pie(
            status_counts,
            values='Quantidade',
            names='Status',
            title='Distribuição de Fases por Status',
            color='Status',
            color_discrete_map={
                'Concluída': '#4CAF50',
                'No Prazo': '#2196F3',
                'Atrasada': '#FF5722',
                'Horas Excedidas': '#FFC107',
                'Não Iniciada': '#9E9E9E'
            }
        )
        
        st.plotly_chart(fig_status, use_container_width=True)