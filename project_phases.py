# project_phases.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from project_phase_manager import ProjectPhaseManager
from database_manager import DatabaseManager, ProjectManager, ClientManager

def format_currency(value):
    """Formata valor monetário"""
    return f"€{value:,.2f}"

def format_hours(hours):
    """Formata horas com duas casas decimais"""
    return f"{hours:.2f}h"

def calcular_proporcao_valores(projeto_info, fase_start_date, fase_end_date):
    """
    Calcula a proporção de valores (horas e custo) que deveria ser alocada para uma fase
    com base na duração da fase em relação à duração total do projeto
    """
    # Converter datas para datetime
    projeto_start = pd.to_datetime(projeto_info['start_date'])
    projeto_end = pd.to_datetime(projeto_info['end_date'])
    fase_start = pd.to_datetime(fase_start_date)
    fase_end = pd.to_datetime(fase_end_date)
    
    # Calcular durações em dias
    duracao_projeto = (projeto_end - projeto_start).days
    duracao_fase = (fase_end - fase_start).days
    
    # Calcular proporção (evita divisão por zero)
    if duracao_projeto <= 0:
        proporcao = 0
    else:
        proporcao = duracao_fase / duracao_projeto
    
    # Calcular valores sugeridos
    horas_sugeridas = float(projeto_info['total_hours']) * proporcao
    custo_sugerido = float(projeto_info['total_cost']) * proporcao
    
    return horas_sugeridas, custo_sugerido

def project_phases_page():
    """Página de gerenciamento de fases de projetos"""
    st.title("Gerenciamento de Fases de Projetos")
    
    # Inicializar gerenciadores
    phase_manager = ProjectPhaseManager()
    db_manager = DatabaseManager()
    project_manager = ProjectManager()
    client_manager = ClientManager()
    
    # Carregar dados dos projetos e clientes
    projects_df = project_manager.read()
    clients_df = client_manager.read()
    
    # Filtrar apenas projetos ativos
    active_projects = projects_df[projects_df['status'] == 'active']
    
    if active_projects.empty:
        st.warning("Não há projetos ativos disponíveis para gerenciamento de fases.")
        return
    
    # Juntar com dados de cliente para exibição
    projects_with_clients = active_projects.merge(
        clients_df[['client_id', 'name']],
        on='client_id',
        how='left'
    )
    
    # Opções para seleção de projeto
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
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Cliente:** {project_info['name']}")
            st.markdown(f"**Período:** {pd.to_datetime(project_info['start_date']).strftime('%d/%m/%Y')} a {pd.to_datetime(project_info['end_date']).strftime('%d/%m/%Y')}")
        with col2:
            st.markdown(f"**Horas Planejadas:** {format_hours(project_info['total_hours'])}")
            st.markdown(f"**Custo Planejado:** {format_currency(project_info['total_cost'])}")
        
        # Obter fases atuais do projeto
        phases = phase_manager.read(project_id=project_id)
        
        # Obter resumo das fases (total de horas e custo já alocados)
        phases_summary = phase_manager.get_project_phases_summary(project_id)
        
        # Calcular recursos disponíveis
        available_resources = phase_manager.get_available_resources(project_id)
        
        if available_resources:
            st.markdown("---")
            st.subheader("Recursos do Projeto")
            
            # Mostrar barras de progresso para horas e custo
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Alocação de Horas**")
                hours_allocated_pct = (phases_summary['total_hours'] / project_info['total_hours']) * 100 if project_info['total_hours'] > 0 else 0
                st.progress(min(hours_allocated_pct / 100, 1.0))
                st.markdown(f"{format_hours(phases_summary['total_hours'])} / {format_hours(project_info['total_hours'])} ({hours_allocated_pct:.1f}%)")
                st.markdown(f"**Disponível:** {format_hours(available_resources['available_hours'])}")
            
            with col2:
                st.markdown("**Alocação de Custo**")
                cost_allocated_pct = (phases_summary['total_cost'] / project_info['total_cost']) * 100 if project_info['total_cost'] > 0 else 0
                st.progress(min(cost_allocated_pct / 100, 1.0))
                st.markdown(f"{format_currency(phases_summary['total_cost'])} / {format_currency(project_info['total_cost'])} ({cost_allocated_pct:.1f}%)")
                st.markdown(f"**Disponível:** {format_currency(available_resources['available_cost'])}")
        
        # Tabela de fases existentes
        st.markdown("---")
        st.subheader("Fases do Projeto")
        
        if not phases.empty:
            # Converter datas para exibição
            display_phases = phases.copy()
            display_phases['start_date'] = pd.to_datetime(display_phases['start_date']).dt.strftime('%d/%m/%Y')
            display_phases['end_date'] = pd.to_datetime(display_phases['end_date']).dt.strftime('%d/%m/%Y')
            
            # Formatar valores
            display_phases['total_hours'] = display_phases['total_hours'].apply(format_hours)
            display_phases['total_cost'] = display_phases['total_cost'].apply(format_currency)
            
            # Exibir tabela de fases
            st.dataframe(
                display_phases[['phase_name', 'start_date', 'end_date', 'total_hours', 'total_cost', 'status']].rename(columns={
                    'phase_name': 'Nome da Fase',
                    'start_date': 'Data Início',
                    'end_date': 'Data Término',
                    'total_hours': 'Horas Planejadas',
                    'total_cost': 'Custo Planejado',
                    'status': 'Status'
                })
            )
            
            # Timeline das fases
            if len(phases) > 0:
                st.subheader("Timeline das Fases")
                
                # Criar dados para o gráfico de Gantt
                gantt_data = []
                
                for _, phase in phases.iterrows():
                    gantt_data.append({
                        'Fase': phase['phase_name'],
                        'Início': pd.to_datetime(phase['start_date']),
                        'Fim': pd.to_datetime(phase['end_date']),
                        'Horas': float(phase['total_hours']),
                        'Custo': float(phase['total_cost'])
                    })
                
                gantt_df = pd.DataFrame(gantt_data)
                
                # Criar gráfico de Gantt
                fig = px.timeline(
                    gantt_df, 
                    x_start='Início', 
                    x_end='Fim', 
                    y='Fase',
                    color='Fase',
                    title='Timeline de Fases',
                    labels={'Fase': 'Nome da Fase'},
                    hover_data=['Horas', 'Custo']
                )
                
                # Adicionar linhas de início e fim do projeto
                fig.add_shape(
                    type='line',
                    x0=pd.to_datetime(project_info['start_date']),
                    y0=-0.5,
                    x1=pd.to_datetime(project_info['start_date']),
                    y1=len(phases) - 0.5,
                    line=dict(color='red', width=2, dash='dash'),
                    name='Início do Projeto'
                )
                
                fig.add_shape(
                    type='line',
                    x0=pd.to_datetime(project_info['end_date']),
                    y0=-0.5,
                    x1=pd.to_datetime(project_info['end_date']),
                    y1=len(phases) - 0.5,
                    line=dict(color='red', width=2, dash='dash'),
                    name='Fim do Projeto'
                )
                
                # Ajustar layout
                fig.update_layout(
                    autosize=True,
                    height=400,
                    xaxis_title='Data',
                    yaxis_title='Fase',
                    xaxis_type='date'
                )
                
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Este projeto ainda não possui fases definidas.")
        
        # Ações para as fases
        st.markdown("---")
        phase_action = st.radio(
            "Ação",
            ["Criar Nova Fase", "Editar Fase", "Excluir Fase"],
            key="phase_action"
        )
        
        if phase_action == "Criar Nova Fase":
            with st.form("create_phase_form"):
                st.subheader("Nova Fase")
                
                # Campos do formulário
                phase_name = st.text_input("Nome da Fase", key="new_phase_name")
                phase_description = st.text_area("Descrição da Fase", key="new_phase_desc")
                
                col1, col2 = st.columns(2)
                with col1:
                    # Usar data de início do projeto como valor padrão
                    default_start = pd.to_datetime(project_info['start_date']).date()
                    phase_start_date = st.date_input("Data de Início", value=default_start, key="new_phase_start")
                
                with col2:
                    # Usar data de fim do projeto como valor padrão
                    default_end = pd.to_datetime(project_info['end_date']).date()
                    phase_end_date = st.date_input("Data de Término", value=default_end, key="new_phase_end")
                
                # Calcular valores sugeridos com base na proporção de tempo
                horas_sugeridas, custo_sugerido = calcular_proporcao_valores(
                    project_info, 
                    phase_start_date, 
                    phase_end_date
                )
                
                # Limitar valores pelas horas e custo disponíveis
                if available_resources:
                    horas_sugeridas = min(horas_sugeridas, available_resources['available_hours'])
                    custo_sugerido = min(custo_sugerido, available_resources['available_cost'])
                
                col3, col4 = st.columns(2)
                with col3:
                    phase_total_hours = st.number_input(
                        "Horas Planejadas",
                        min_value=0.0,
                        max_value=float(available_resources['available_hours']) if available_resources else float(project_info['total_hours']),
                        value=float(horas_sugeridas),
                        step=0.5,
                        format="%.2f",
                        key="new_phase_hours"
                    )
                
                with col4:
                    phase_total_cost = st.number_input(
                        "Custo Planejado (€)",
                        min_value=0.0,
                        max_value=float(available_resources['available_cost']) if available_resources else float(project_info['total_cost']),
                        value=float(custo_sugerido),
                        step=10.0,
                        format="%.2f",
                        key="new_phase_cost"
                    )
                
                phase_status = st.selectbox(
                    "Status",
                    options=["active", "pending", "completed"],
                    index=0,
                    key="new_phase_status"
                )
                
                # Botão para criar fase
                submit_button = st.form_submit_button("Criar Fase")
                
                if submit_button:
                    # Validar datas
                    if phase_start_date > phase_end_date:
                        st.error("A data de término deve ser posterior à data de início.")
                    # Validar nome
                    elif not phase_name:
                        st.error("O nome da fase é obrigatório.")
                    # Validar datas do projeto
                    elif phase_start_date < pd.to_datetime(project_info['start_date']).date() or phase_end_date > pd.to_datetime(project_info['end_date']).date():
                        st.error("As datas da fase devem estar dentro do período do projeto.")
                    else:
                        # Criar objeto de dados
                        phase_data = {
                            'project_id': project_id,
                            'phase_name': phase_name,
                            'phase_description': phase_description,
                            'start_date': phase_start_date.isoformat(),
                            'end_date': phase_end_date.isoformat(),
                            'total_hours': phase_total_hours,
                            'total_cost': phase_total_cost,
                            'status': phase_status
                        }
                        
                        # Tentar criar a fase
                        success, result = phase_manager.create(phase_data)
                        
                        if success:
                            st.success(f"Fase '{phase_name}' criada com sucesso!")
                            # Recarregar página para mostrar a nova fase
                            st.rerun()
                        else:
                            st.error(f"Erro ao criar fase: {result}")
        
        elif phase_action == "Editar Fase" and not phases.empty:
            # Opções para seleção de fase
            phase_options = []
            for _, phase in phases.iterrows():
                phase_options.append({
                    'id': phase['phase_id'],
                    'display': f"{phase['phase_name']} ({pd.to_datetime(phase['start_date']).strftime('%d/%m/%Y')} a {pd.to_datetime(phase['end_date']).strftime('%d/%m/%Y')})"
                })
            
            # Seleção de fase para editar
            selected_phase = st.selectbox(
                "Selecione a Fase para Editar",
                options=phase_options,
                format_func=lambda x: x['display']
            )
            
            if selected_phase:
                phase_id = selected_phase['id']
                phase_info = phases[phases['phase_id'] == phase_id].iloc[0]
                
                # Recursos disponíveis considerando a fase atual
                available_with_current = phase_manager.get_available_resources(project_id, phase_id)
                
                # Formulário de edição
                with st.form("edit_phase_form"):
                    st.subheader("Editar Fase")
                    
                    # Campos do formulário com valores atuais
                    phase_name = st.text_input("Nome da Fase", value=phase_info['phase_name'], key="edit_phase_name")
                    phase_description = st.text_area("Descrição da Fase", value=phase_info['phase_description'] if pd.notna(phase_info['phase_description']) else "", key="edit_phase_desc")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        phase_start_date = st.date_input(
                            "Data de Início", 
                            value=pd.to_datetime(phase_info['start_date']).date(),
                            key="edit_phase_start"
                        )
                    
                    with col2:
                        phase_end_date = st.date_input(
                            "Data de Término", 
                            value=pd.to_datetime(phase_info['end_date']).date(),
                            key="edit_phase_end"
                        )
                    
                    col3, col4 = st.columns(2)
                    with col3:
                        # Calcular máximo de horas disponíveis (horas da fase atual + horas disponíveis no projeto)
                        max_hours = float(phase_info['total_hours']) + float(available_with_current['available_hours']) if available_with_current else float(project_info['total_hours'])
                        
                        phase_total_hours = st.number_input(
                            "Horas Planejadas",
                            min_value=0.0,
                            max_value=max_hours,
                            value=float(phase_info['total_hours']),
                            step=0.5,
                            format="%.2f",
                            key="edit_phase_hours"
                        )
                    
                    with col4:
                        # Calcular máximo de custo disponível (custo da fase atual + custo disponível no projeto)
                        max_cost = float(phase_info['total_cost']) + float(available_with_current['available_cost']) if available_with_current else float(project_info['total_cost'])
                        
                        phase_total_cost = st.number_input(
                            "Custo Planejado (€)",
                            min_value=0.0,
                            max_value=max_cost,
                            value=float(phase_info['total_cost']),
                            step=10.0,
                            format="%.2f",
                            key="edit_phase_cost"
                        )
                    
                    phase_status = st.selectbox(
                        "Status",
                        options=["active", "pending", "completed"],
                        index=["active", "pending", "completed"].index(phase_info['status']),
                        key="edit_phase_status"
                    )
                    
                    # Botão para atualizar fase
                    submit_button = st.form_submit_button("Atualizar Fase")
                    
                    if submit_button:
                        # Validar datas
                        if phase_start_date > phase_end_date:
                            st.error("A data de término deve ser posterior à data de início.")
                        # Validar nome
                        elif not phase_name:
                            st.error("O nome da fase é obrigatório.")
                        # Validar datas do projeto
                        elif phase_start_date < pd.to_datetime(project_info['start_date']).date() or phase_end_date > pd.to_datetime(project_info['end_date']).date():
                            st.error("As datas da fase devem estar dentro do período do projeto.")
                        else:
                            # Criar objeto de dados
                            phase_data = {
                                'phase_name': phase_name,
                                'phase_description': phase_description,
                                'start_date': phase_start_date.isoformat(),
                                'end_date': phase_end_date.isoformat(),
                                'total_hours': phase_total_hours,
                                'total_cost': phase_total_cost,
                                'status': phase_status
                            }
                            
                            # Tentar atualizar a fase
                            success, result = phase_manager.update(phase_id, phase_data)
                            
                            if success:
                                st.success(f"Fase '{phase_name}' atualizada com sucesso!")
                                # Recarregar página para mostrar a fase atualizada
                                st.rerun()
                            else:
                                st.error(f"Erro ao atualizar fase: {result}")
        
        elif phase_action == "Excluir Fase" and not phases.empty:
            # Opções para seleção de fase
            phase_options = []
            for _, phase in phases.iterrows():
                phase_options.append({
                    'id': phase['phase_id'],
                    'display': f"{phase['phase_name']} ({pd.to_datetime(phase['start_date']).strftime('%d/%m/%Y')} a {pd.to_datetime(phase['end_date']).strftime('%d/%m/%Y')})"
                })
            
            # Seleção de fase para excluir
            selected_phase = st.selectbox(
                "Selecione a Fase para Excluir",
                options=phase_options,
                format_func=lambda x: x['display']
            )
            
            if selected_phase:
                phase_id = selected_phase['id']
                phase_info = phases[phases['phase_id'] == phase_id].iloc[0]
                
                # Confirmação de exclusão
                st.warning(f"Você está prestes a excluir a fase '{phase_info['phase_name']}'. Esta ação não pode ser desfeita.")
                
                if st.button("Confirmar Exclusão", key="confirm_delete_phase"):
                    # Excluir fase
                    if phase_manager.delete(phase_id):
                        st.success(f"Fase '{phase_info['phase_name']}' excluída com sucesso!")
                        # Recarregar página
                        st.rerun()
                    else:
                        st.error("Erro ao excluir fase. Por favor, tente novamente.")

def integrate_phases_with_project_reports(project_id):
    """
    Integra informações das fases nos relatórios de projetos existentes
    """
    phase_manager = ProjectPhaseManager()
    
    # Obter fases do projeto
    phases = phase_manager.read(project_id=project_id)
    
    if phases.empty:
        return None
    
    # Resumo das fases
    phases_summary = {
        'total_phases': len(phases),
        'phases': []
    }
    
    # Processar cada fase
    for _, phase in phases.iterrows():
        phases_summary['phases'].append({
            'name': phase['phase_name'],
            'start_date': pd.to_datetime(phase['start_date']),
            'end_date': pd.to_datetime(phase['end_date']),
            'total_hours': float(phase['total_hours']),
            'total_cost': float(phase['total_cost']),
            'status': phase['status']
        })
    
    return phases_summary