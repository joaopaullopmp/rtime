import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from timesheet import TimesheetManager
from report_utils import calcular_dias_uteis_projeto
from database_manager import DatabaseManager
from project_phases import integrate_phases_with_project_reports
import calendar
from project_report_button import add_report_export_button

def format_hours_minutes(hours):
    """Converte horas decimais para formato HH:mm"""
    total_minutes = int(hours * 60)
    hours_part = total_minutes // 60
    minutes_part = total_minutes % 60
    return f"{hours_part:02d}:{minutes_part:02d}"

def reports_page():
    """Página principal de relatórios de projetos"""
    st.title("Relatório de Projetos")
    
    try:
        # Inicializar gerenciadores
        db_manager = DatabaseManager()
        timesheet = TimesheetManager()
        
        # Carregar dados das tabelas
        clients_df = db_manager.query_to_df("SELECT * FROM clients")
        projects_df = db_manager.query_to_df("SELECT * FROM projects")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        rates_df = db_manager.query_to_df("SELECT * FROM rates")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        
        # Verificar o papel do usuário atual e obter informações relevantes
        is_admin = st.session_state.user_info['role'].lower() == 'admin'
        is_leader = st.session_state.user_info['role'].lower() == 'leader'
        current_user_id = st.session_state.user_info['user_id']
        
        # Para líderes, identificar sua equipe
        leader_team = None
        leader_team_id = None
        
        if is_leader:
            user_data = users_df[users_df['user_id'] == current_user_id]
            
            if not user_data.empty:
                # Obter a equipe do líder
                user_groups = eval(user_data['groups'].iloc[0]) if isinstance(user_data['groups'].iloc[0], str) else []
                if isinstance(user_groups, dict):
                    user_groups = list(user_groups.values())
                
                leader_team = user_groups[0] if user_groups else None
                
                if leader_team:
                    st.info(f"Visualizando projetos da equipe: {leader_team}")
                    
                    # Buscar ID da equipe
                    team_row = groups_df[groups_df['group_name'] == leader_team]
                    if not team_row.empty:
                        leader_team_id = team_row['id'].iloc[0]
                       
        # Filtros por equipa e tipo de projeto
         
        with st.expander("Filtros", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                if is_admin:
                    all_groups = groups_df[groups_df['active'] == True]['group_name'].tolist()
                    selected_group = st.selectbox(
                        "Filtrar por Equipa",
                        options=["Todas"] + all_groups,
                        key="group_filter"
                    )
                else:
                    # Para líderes, mostrar apenas sua equipe
                    selected_group = leader_team if leader_team else "Sem equipe"
                    st.text(f"Equipa: {selected_group}")
            
            with col2:
                # Permitir seleção múltipla de tipos de projeto
                project_types = projects_df['project_type'].dropna().unique().tolist()
                selected_types = st.multiselect(
                    "Filtrar por Tipo de Projeto",
                    options=project_types,
                    default=[],
                    key="type_filter"
                )
        
        # Aplicar filtros aos clientes
        active_clients = clients_df[clients_df['active'] == True].copy()
        
        # Para líderes, filtrar apenas os clientes da sua equipe
        if is_leader and leader_team_id is not None:
            active_clients = active_clients[active_clients['group_id'] == leader_team_id]
        
        # Aplicar filtro de tipo de projeto aos clientes
        if selected_types:
            # Encontrar projetos dos tipos selecionados
            filtered_projects = projects_df[projects_df['project_type'].isin(selected_types)]
            # Obter client_ids dos projetos filtrados
            client_ids_with_selected_types = filtered_projects['client_id'].unique()
            # Filtrar clientes que tenham pelo menos um projeto dos tipos selecionados
            active_clients = active_clients[active_clients['client_id'].isin(client_ids_with_selected_types)]
        
        # Aplicar filtro de grupo para admin
        if is_admin and selected_group != "Todas":
            group_id = groups_df[groups_df['group_name'] == selected_group]['id'].iloc[0]
            active_clients = active_clients[active_clients['group_id'] == group_id]
        
        if active_clients.empty:
            if is_leader:
                st.warning(f"Não há clientes disponíveis para a sua equipe.")
            else:
                filtro_texto = f" do tipo {', '.join(selected_types)}" if selected_types else ""
                st.warning(f"Não há clientes disponíveis para a equipa '{selected_group}'{filtro_texto}.")
            return
        
        # Seleção de Cliente
        client = st.selectbox(
            "Cliente",
            options=active_clients['name'].tolist()
        )

        if client:
            client_id = active_clients[active_clients['name'] == client]['client_id'].iloc[0]
            
            # Aplicar filtro de tipo de projeto (múltiplos tipos)
            client_projects = projects_df[projects_df['client_id'] == client_id]
            if selected_types:
                client_projects = client_projects[client_projects['project_type'].isin(selected_types)]
            
            # Para líderes, garantir que só veem projetos da sua equipe
            if is_leader and leader_team_id is not None:
                client_projects = client_projects[client_projects['group_id'] == leader_team_id]
            
            if client_projects.empty:
                if is_leader:
                    st.warning(f"Não há projetos disponíveis para este cliente na sua equipe.")
                else:
                    tipos_txt = ", ".join(selected_types) if selected_types else "selecionados" 
                    st.warning(f"Não há projetos dos tipos {tipos_txt} para o cliente '{client}'.")
                return
            
            projeto = st.selectbox(
                "Projeto",
                options=client_projects['project_name'].tolist()
            )

            if projeto:
                try:
                    projeto_info = projects_df[projects_df['project_name'] == projeto].iloc[0]
                    
                    # Verificar dados essenciais
                    missing_fields = []
                    if pd.isna(projeto_info['total_cost']):
                        missing_fields.append("Custo Total")
                    if pd.isna(projeto_info['total_hours']):
                        missing_fields.append("Horas Totais")
                        
                    if missing_fields:
                        st.warning(f"Este projeto não possui os seguintes dados essenciais: {', '.join(missing_fields)}. Alguns cálculos podem estar incorretos.")
                    
                    # Carregar entradas de timesheet
                    projeto_id = int(projeto_info['project_id'])
                    entries = db_manager.query_to_df(f"SELECT * FROM timesheet WHERE project_id = {projeto_id}")
                    
                    # Verificar e tratar horas migradas (horas_realizadas_mig)
                    horas_migradas = 0
                    if 'horas_realizadas_mig' in projeto_info and not pd.isna(projeto_info['horas_realizadas_mig']):
                        horas_migradas = float(projeto_info['horas_realizadas_mig'])
                    
                    # Verificar e tratar custo migrado (custo_realizado_mig)
                    custo_migrado = 0
                    if 'custo_realizado_mig' in projeto_info and not pd.isna(projeto_info['custo_realizado_mig']):
                        custo_migrado = float(projeto_info['custo_realizado_mig'])
                    
                    # Calcular horas regulares e extras
                    horas_regulares = 0
                    horas_extras = 0
                    custo_extras = 0
                    
                    if not entries.empty:
                        # Separar horas extras e regulares
                        horas_regulares = entries[~entries['overtime'].astype(bool)]['hours'].sum()
                        # Para horas extras, vamos guardar o valor original para informação, mas contabilizar o dobro
                        horas_extras_originais = entries[entries['overtime'].astype(bool)]['hours'].sum()
                        # Calculamos o dobro para contabilização
                        horas_extras = horas_extras_originais * 2
                    
                    # Total de horas (regulares + extras*2)
                    horas_realizadas = float(horas_regulares + horas_extras)
                    
                    # CORREÇÃO: Adicionar horas migradas ao total
                    horas_realizadas_total = horas_realizadas + horas_migradas
                    
                    # Calcular custo realizado
                    custo_realizado = 0
                    custo_horas_extras = 0
                    if not entries.empty:
                        for _, entry in entries.iterrows():
                            try:
                                user_id = entry['user_id']
                                hours = float(entry['hours'])
                                is_overtime = entry['overtime'] if 'overtime' in entry else False
                                # Converter para booleano caso seja outro tipo de dado
                                if isinstance(is_overtime, (int, float)):
                                    is_overtime = bool(is_overtime)
                                elif isinstance(is_overtime, str):
                                    is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                                
                                # Obter rate para o usuário
                                rate_value = None
                                if 'rate_value' in entry and not pd.isna(entry['rate_value']):
                                    rate_value = float(entry['rate_value'])
                                else:
                                    user_info = db_manager.query_to_df(f"SELECT rate_id FROM utilizadores WHERE user_id = {user_id}")
                                    
                                    if not user_info.empty and not pd.isna(user_info['rate_id'].iloc[0]):
                                        rate_id = user_info['rate_id'].iloc[0]
                                        rate_info = db_manager.query_to_df(f"SELECT rate_cost FROM rates WHERE rate_id = {rate_id}")
                                        
                                        if not rate_info.empty:
                                            rate_value = float(rate_info['rate_cost'].iloc[0])
                                
                                # Calcular custo com base no rate obtido
                                if rate_value:
                                    entry_cost = hours * rate_value
                                    
                                    # Se for hora extra, multiplicar por 2
                                    if is_overtime:
                                        custo_horas_extras += entry_cost
                                        custo_realizado += entry_cost * 2  # Dobro para horas extras
                                    else:
                                        custo_realizado += entry_cost  # Normal para horas regulares
                            except Exception as e:
                                debug_flag = False  # Definir como True para depuração
                                if debug_flag:
                                    st.sidebar.write(f"Erro ao processar entrada: {e}")
                    
                    # CORREÇÃO: Adicionar custo migrado ao total
                    custo_realizado_total = custo_realizado + custo_migrado
                    
                    # Garantir valores padrão
                    projeto_info['total_cost'] = float(projeto_info['total_cost']) if pd.notna(projeto_info['total_cost']) else 0.0
                    projeto_info['total_hours'] = float(projeto_info['total_hours']) if pd.notna(projeto_info['total_hours']) else 0.0
                    
                    # Calcular métricas baseadas em dias úteis
                    data_inicio = pd.to_datetime(projeto_info['start_date']).date()
                    data_fim = pd.to_datetime(projeto_info['end_date']).date()
                    data_atual = min(datetime.now().date(), data_fim)
                    
                    dias_uteis_totais = calcular_dias_uteis_projeto(data_inicio, data_fim)
                    dias_uteis_decorridos = calcular_dias_uteis_projeto(data_inicio, data_atual)
                    dias_uteis_restantes = dias_uteis_totais - dias_uteis_decorridos
                    
                    percentual_tempo_decorrido = dias_uteis_decorridos / dias_uteis_totais if dias_uteis_totais > 0 else 0
                    horas_diarias_planejadas = projeto_info['total_hours'] / dias_uteis_totais if dias_uteis_totais > 0 else 0
                    horas_planejadas_ate_agora = horas_diarias_planejadas * dias_uteis_decorridos
                    custo_planejado_ate_agora = projeto_info['total_cost'] * percentual_tempo_decorrido
                    
                    # CORREÇÃO: Usar valores totais (incluindo migrados) para CPI
                    cpi = custo_planejado_ate_agora / custo_realizado_total if custo_realizado_total > 0 else 1.0
                    
                    # CORREÇÃO: Usar valores totais para EAC e VAC
                    eac = custo_realizado_total
                    if dias_uteis_decorridos > 0 and cpi != 0:
                        custo_diario_real = custo_realizado_total / dias_uteis_decorridos
                        custo_projetado_restante = (custo_diario_real / cpi) * dias_uteis_restantes
                        eac = custo_realizado_total + custo_projetado_restante
                    vac = projeto_info['total_cost'] - eac
                    
                    # Armazenar métricas em um dicionário
                    metricas = {
                        'cpi': cpi,
                        'custo_planejado': custo_planejado_ate_agora,
                        'custo_realizado': custo_realizado_total,  # Já inclui custo migrado
                        'custo_realizado_atual': custo_realizado,  # Apenas custo atual
                        'custo_realizado_migrado': custo_migrado,  # Apenas custo migrado
                        'custo_horas_extras': custo_horas_extras,  # Custo original das horas extras
                        'eac': eac,
                        'vac': vac,
                        'horas_realizadas': horas_realizadas_total,  # Já inclui horas migradas
                        'horas_realizadas_atual': horas_realizadas,  # Apenas horas atuais
                        'horas_realizadas_migrado': horas_migradas,  # Apenas horas migradas
                        'horas_regulares': horas_regulares,  # Apenas horas regulares
                        'horas_extras_originais': horas_extras_originais if 'horas_extras_originais' in locals() else 0,  # Horas extras antes da multiplicação
                        'horas_extras': horas_extras,  # Horas extras já multiplicadas por 2
                        'dias_uteis_totais': dias_uteis_totais,
                        'dias_uteis_decorridos': dias_uteis_decorridos,
                        'dias_uteis_restantes': dias_uteis_restantes,
                        'horas_diarias_planejadas': horas_diarias_planejadas,
                        'horas_planejadas_ate_agora': horas_planejadas_ate_agora
                    }
                    
                except Exception as e:
                    st.error(f"Erro ao carregar informações do projeto: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())
                    return

                # 1. Card Principal com Informações Básicas
                st.markdown("### 📑 Informações do Projeto")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Cliente", client)
                    st.metric("Projeto", projeto_info['project_name'])
                    # ADICIONADO: Tipo de projeto
                    st.metric("Tipo de Projeto", projeto_info['project_type'])
                with col2:
                    start_date = pd.to_datetime(projeto_info['start_date']).strftime('%d/%m/%Y')
                    end_date = pd.to_datetime(projeto_info['end_date']).strftime('%d/%m/%Y')
                    st.metric("Data Início", start_date)
                    st.metric("Data Término", end_date)
                with col3:
                    st.metric("Status", projeto_info['status'].title())
                    
                    # MODIFICADO: Colorir o CPI baseado no valor (vermelho se < 1, verde se >= 1)
                    # Definir cor do texto do CPI
                    cpi_color = "green" if metricas['cpi'] >= 1 else "red"
                    delta_color = "normal" if metricas['cpi'] >= 1 else "inverse"
                    
                    # Criar um HTML com a cor aplicada ao valor do CPI
                    cpi_html = f"""
                    <div style="display: flex; flex-direction: column;">
                        <div style="font-size: 14px; color: #31333F;">CPI</div>
                        <div style="font-size: 24px; font-weight: 700; color: {cpi_color};">{metricas['cpi']:.2f}</div>
                        <div style="font-size: 12px; color: #6E7884;">CPI >= 1: Projeto no orçamento<br>CPI < 1: Projeto acima do orçamento</div>
                    </div>
                    """
                    st.markdown(cpi_html, unsafe_allow_html=True)

                st.markdown("---")
                
                # 2. Dashboard Principal
                st.subheader("📊 Dashboard do Projeto")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(
                        "Custo Total Planejado", 
                        f"€{projeto_info['total_cost']:,.2f}",
                        help="Custo total planejado para o projeto"
                    )
                    
                    # Exibir horas no formato HH:mm
                    horas_planejadas_fmt = format_hours_minutes(projeto_info['total_hours'])
                    horas_realizadas_fmt = format_hours_minutes(metricas['horas_realizadas'])
                    
                    # Adicionar informação sobre horas extras
                    total_horas_info = f"{horas_realizadas_fmt} realizadas"
                    if metricas['horas_extras_originais'] > 0:
                        horas_extras_fmt = format_hours_minutes(metricas['horas_extras_originais'])
                        total_horas_info = f"{horas_realizadas_fmt} realizadas (inclui {horas_extras_fmt} em horas extras)"
                    
                    st.metric(
                        "Total de Horas Planejadas", 
                        f"{horas_planejadas_fmt}",
                        total_horas_info
                    )
                
                with col2:
                    delta_cor = "inverse" if metricas['custo_realizado'] > metricas['custo_planejado'] else "normal"
                    st.metric(
                        "Custo Realizado", 
                        f"€{metricas['custo_realizado']:,.2f}",
                        f"€{metricas['custo_planejado']:,.2f} planejado até agora",
                        delta_color=delta_cor
                    )
                    
                    # Exibir horas disponíveis no formato HH:mm
                    horas_disponiveis = projeto_info['total_hours'] - metricas['horas_realizadas']
                    horas_disponiveis_fmt = format_hours_minutes(max(0, horas_disponiveis))  # Garantir não negativo
                    
                    st.metric(
                        "Horas Disponíveis", 
                        f"{horas_disponiveis_fmt}",
                        help="Total de Horas disponíveis no projeto"
                    )
                
                with col3:
                    delta_cor = "inverse" if metricas['eac'] > projeto_info['total_cost'] else "normal"
                    st.metric(
                        "Projeção Final (EAC)", 
                        f"€{metricas['eac']:,.2f}",
                        f"€{metricas['vac']:,.2f} vs orçamento",
                        delta_color=delta_cor
                    )
                    
                    if projeto_info['total_hours'] > 0:
                        perc_horas = (metricas['horas_realizadas'] / projeto_info['total_hours']) * 100
                        st.metric("% Horas Consumidas", f"{perc_horas:.1f}%")
                        st.progress(min(perc_horas/100, 1.0))
                
                st.markdown("---")
                
                # NOVA SEÇÃO: Consumo de Horas Mensais (Planejadas vs. Realizadas)
                st.subheader("📈 Consumo de Horas Mensais")
                
                # Calcular dados de consumo mensal
                try:
                    # Período do projeto
                    data_inicio_projeto = pd.to_datetime(projeto_info['start_date'], format='mixed')
                    data_fim_projeto = pd.to_datetime(projeto_info['end_date'], format='mixed')
                    
                    # Criar uma lista de meses entre o início e fim do projeto
                    meses = []
                    atual = data_inicio_projeto.replace(day=1)
                    while atual <= data_fim_projeto:
                        meses.append(atual)
                        # Avançar para o próximo mês
                        if atual.month == 12:
                            atual = atual.replace(year=atual.year + 1, month=1)
                        else:
                            atual = atual.replace(month=atual.month + 1)
                    
                    # Calcular horas planejadas por mês (distribuição uniforme)
                    total_meses = len(meses)
                    horas_por_mes_planejadas = projeto_info['total_hours'] / total_meses if total_meses > 0 else 0
                    
                    # Preparar dados para o gráfico de consumo mensal
                    consumo_mensal = []
                    
                    for mes in meses:
                        # Determinar o último dia do mês
                        ultimo_dia_mes = calendar.monthrange(mes.year, mes.month)[1]
                        fim_mes = mes.replace(day=ultimo_dia_mes)
                        
                        # Calcular dias úteis no mês
                        dias_uteis_mes = calcular_dias_uteis_projeto(mes.date(), fim_mes.date())
                        
                        # Horas planejadas para o mês (pode ser refinado com base em dias úteis)
                        horas_planejadas_mes = (dias_uteis_mes / dias_uteis_totais) * projeto_info['total_hours'] if dias_uteis_totais > 0 else 0
                        
                        # Filtrar entradas de timesheet para o mês
                        if not entries.empty:
                            # Converter start_date para datetime
                            entries['start_date_dt'] = pd.to_datetime(entries['start_date'], format='mixed')
    
                            # Use the converted datetime column for filtering
                            entries_mes = entries[
                                (entries['start_date_dt'].dt.year == mes.year) &
                                (entries['start_date_dt'].dt.month == mes.month)
    ]
                            
                            # Calcular horas registradas no mês
                            horas_regulares_mes = entries_mes[~entries_mes['overtime'].astype(bool)]['hours'].sum() if not entries_mes.empty else 0
                            horas_extras_mes_orig = entries_mes[entries_mes['overtime'].astype(bool)]['hours'].sum() if not entries_mes.empty else 0
                            horas_extras_mes = horas_extras_mes_orig * 2
                            horas_realizadas_mes = horas_regulares_mes + horas_extras_mes
                        else:
                            horas_realizadas_mes = 0
                            horas_regulares_mes = 0
                            horas_extras_mes_orig = 0
                            horas_extras_mes = 0
                        
                        # Tratar meses passados vs. futuros
                        data_atual = datetime.now()
                        
                        # Adicionar horas migradas apenas para o primeiro mês
                        horas_migradas_mes = horas_migradas if mes == meses[0] else 0
                        
                        if mes.year < data_atual.year or (mes.year == data_atual.year and mes.month < data_atual.month):
                            # Mês passado - usar dados reais
                            status = "Passado"
                        elif mes.year == data_atual.year and mes.month == data_atual.month:
                            # Mês atual - em andamento
                            status = "Atual"
                        else:
                            # Mês futuro - planejado
                            status = "Futuro"
                        
                        # Calcular percentual em relação ao planejado
                        horas_realizadas_total_mes = horas_realizadas_mes + horas_migradas_mes
                        percentual = (horas_realizadas_total_mes / horas_planejadas_mes * 100) if horas_planejadas_mes > 0 else 0
                        
                        # Adicionar ao array de consumo mensal
                        consumo_mensal.append({
                            'mes': mes,
                            'mes_str': mes.strftime('%b/%Y'),
                            'horas_planejadas': horas_planejadas_mes,
                            'horas_realizadas': horas_realizadas_total_mes,
                            'horas_regulares': horas_regulares_mes,
                            'horas_extras_orig': horas_extras_mes_orig,
                            'horas_extras': horas_extras_mes,
                            'horas_migradas': horas_migradas_mes,
                            'percentual': percentual,
                            'status': status
                        })
                    
                    # Converter para DataFrame
                    consumo_df = pd.DataFrame(consumo_mensal)
                    
                    # Exibir gráfico de consumo mensal
                    if not consumo_df.empty:
                        # Criar gráfico combinado (barras + linha)
                        fig = go.Figure()
                        
                        # Barras para horas realizadas (empilhadas: regulares, extras, migradas)
                        fig.add_trace(go.Bar(
                            x=consumo_df['mes_str'],
                            y=consumo_df['horas_regulares'],
                            name='Horas Regulares',
                            marker_color='#2196F3'
                        ))
                        
                        fig.add_trace(go.Bar(
                            x=consumo_df['mes_str'],
                            y=consumo_df['horas_extras'],
                            name='Horas Extras (2x)',
                            marker_color='#FF9800'
                        ))
                        
                        fig.add_trace(go.Bar(
                            x=consumo_df['mes_str'],
                            y=consumo_df['horas_migradas'],
                            name='Horas Migradas',
                            marker_color='#9C27B0'
                        ))
                        
                        # Linha para horas planejadas
                        fig.add_trace(go.Scatter(
                            x=consumo_df['mes_str'],
                            y=consumo_df['horas_planejadas'],
                            name='Horas Planejadas',
                            mode='lines+markers',
                            line=dict(color='#4CAF50', width=2, dash='solid')
                        ))
                        
                        # Configurar layout
                        fig.update_layout(
                            title='Consumo de Horas Mensais',
                            xaxis_title='Mês',
                            yaxis_title='Horas',
                            barmode='stack',
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1
                            ),
                            height=500
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Tabela de consumo mensal
                        st.subheader("Detalhamento Mensal")
                        
                        # Preparar dataframe para exibição
                        display_df = consumo_df.copy()
                        
                        # Formatar colunas numéricas para exibição
                        display_columns = {
                            'Mês': 'mes_str',
                            'Status': 'status',
                            'Planejado': 'horas_planejadas',
                            'Realizado': 'horas_realizadas',
                            'Regulares': 'horas_regulares',
                            'Extras (2x)': 'horas_extras',
                            'Migradas': 'horas_migradas',
                            'Percentual': 'percentual'
                        }
                        
                        # Selecionar e renomear colunas
                        display_df = display_df[list(display_columns.values())].rename(columns={v: k for k, v in display_columns.items()})
                        
                        # Formatar colunas numéricas (horas no formato HH:MM e percentual com %)
                        for col in ['Planejado', 'Realizado', 'Regulares', 'Extras (2x)', 'Migradas']:
                            display_df[col] = display_df[col].apply(format_hours_minutes)
                        
                        display_df['Percentual'] = display_df['Percentual'].apply(lambda x: f"{x:.1f}%")
                        
                        # Destacar mês atual
                        def highlight_status(val):
                            if val == 'Atual':
                                return 'background-color: #e3f2fd'
                            elif val == 'Passado':
                                return 'background-color: #f1f8e9'
                            return ''
                        
                        # Exibir tabela formatada
                        st.dataframe(
                            display_df.style.applymap(highlight_status, subset=['Status'])
                        )
                        
                        # Adicionar resumo
                        st.subheader("Resumo do Consumo")
                        
                        # Calcular totais
                        total_planejado = consumo_df['horas_planejadas'].sum()
                        total_realizado = consumo_df['horas_realizadas'].sum()
                        
                        # Calcular média mensal e projeção
                        meses_passados = consumo_df[consumo_df['status'] != 'Futuro']
                        media_mensal = meses_passados['horas_realizadas'].mean() if not meses_passados.empty else 0
                        meses_futuros = len(consumo_df[consumo_df['status'] == 'Futuro'])
                        
                        projecao_total = total_realizado + (media_mensal * meses_futuros)
                        
                        # Exibir métricas
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.metric(
                                "Total Planejado", 
                                format_hours_minutes(total_planejado)
                            )
                        
                        with col2:
                            st.metric(
                                "Total Realizado", 
                                format_hours_minutes(total_realizado),
                                f"{(total_realizado/total_planejado*100):.1f}% do planejado" if total_planejado > 0 else "N/A"
                            )
                        
                        with col3:
                            st.metric(
                                "Projeção Final", 
                                format_hours_minutes(projecao_total),
                                f"{(projecao_total/projeto_info['total_hours']*100):.1f}% do total" if projeto_info['total_hours'] > 0 else "N/A"
                            )
                            
                except Exception as e:
                    st.warning(f"Não foi possível calcular o consumo mensal: {str(e)}")
                    import traceback
                    st.sidebar.error(traceback.format_exc())
                
                st.markdown("---")
                
                # 3. Utilização por Recursos (incluindo dados migrados)
                st.subheader("🧑‍💼 Utilização de Recursos")
                
                # Inicializar dataframe de recursos
                recursos_dados = []
                
                # Adicionar dados do sistema atual
                if not entries.empty:
                    # Mesclar com informações de usuários para obter nomes
                    recursos_df = entries.merge(
                        users_df[['user_id', 'First_Name', 'Last_Name', 'rate_id']],
                        on='user_id',
                        how='left'
                    )
                    
                    # Adicionar informações de custo baseadas nas rates, considerando horas extras
                    def calcular_custo_entrada(row):
                        if pd.isna(row['rate_id']):
                            return 0
                        
                        rate_info = rates_df[rates_df['rate_id'] == row['rate_id']]
                        if rate_info.empty:
                            return 0
                        
                        rate_value = float(rate_info['rate_cost'].iloc[0])
                        is_overtime = row.get('overtime', False)
                        
                        # Converter is_overtime para booleano
                        if isinstance(is_overtime, (int, float)):
                            is_overtime = bool(is_overtime)
                        elif isinstance(is_overtime, str):
                            is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                        
                        # Multiplicar por 2 se for hora extra
                        multiplicador = 2 if is_overtime else 1
                        return float(row['hours']) * rate_value * multiplicador
                    
                    recursos_df['custo'] = recursos_df.apply(calcular_custo_entrada, axis=1)
                    
                    # Coluna para indicar se é hora extra
                    recursos_df['is_extra'] = recursos_df.apply(
                        lambda row: bool(row.get('overtime', False)), 
                        axis=1
                    )
                    
                    # Criar coluna de nome completo
                    recursos_df['nome_completo'] = recursos_df.apply(
                        lambda row: f"{row['First_Name']} {row['Last_Name']}" if pd.notna(row['First_Name']) else "Desconhecido",
                        axis=1
                    )
                    
                    # Agrupar por recurso (usuário)
                    for nome, dados in recursos_df.groupby('nome_completo'):
                        horas_total = dados['hours'].sum()
                        custo_total = dados['custo'].sum()
                        recursos_dados.append({
                            'nome_completo': nome,
                            'hours': horas_total,
                            'custo': custo_total
                        })
                
                # CORREÇÃO: Adicionar dados migrados como recurso "GLPI" se existirem
                if metricas['horas_realizadas_migrado'] > 0 or metricas['custo_realizado_migrado'] > 0:
                    recursos_dados.append({
                        'nome_completo': 'GLPI',
                        'hours': metricas['horas_realizadas_migrado'],
                        'custo': metricas['custo_realizado_migrado']
                    })
                
                # Converter para DataFrame
                recursos_utilizacao = pd.DataFrame(recursos_dados)
                
                if not recursos_utilizacao.empty:
                    # Adicionar percentual sobre o total
                    total_horas = recursos_utilizacao['hours'].sum()
                    total_custo = recursos_utilizacao['custo'].sum()
                    
                    # CORREÇÃO: Verificar se total_horas e total_custo coincidem com os valores esperados
                    # Adicionar aviso sobre horas extras
                    if metricas['horas_extras_originais'] > 0:
                        st.info(f"⚠️ **Atenção**: Este projeto possui **{format_hours_minutes(metricas['horas_extras_originais'])}** de horas extras, que são contabilizadas em dobro tanto em horas quanto em custo.", icon="ℹ️")
                    
                    if abs(total_horas - metricas['horas_realizadas']) > 0.01:
                        st.warning(f"Possível inconsistência: total de horas na tabela ({total_horas:.2f}) difere do total calculado ({metricas['horas_realizadas']:.2f})")
                    
                    if abs(total_custo - metricas['custo_realizado']) > 0.01:
                        st.warning(f"Possível inconsistência: custo total na tabela (€{total_custo:.2f}) difere do custo total calculado (€{metricas['custo_realizado']:.2f})")
                    
                    recursos_utilizacao['perc_horas'] = recursos_utilizacao['hours'].apply(
                        lambda x: (x / total_horas * 100) if total_horas > 0 else 0
                    )
                    recursos_utilizacao['perc_custo'] = recursos_utilizacao['custo'].apply(
                        lambda x: (x / total_custo * 100) if total_custo > 0 else 0
                    )
                    
                    # Ordenar por horas (decrescente)
                    recursos_utilizacao = recursos_utilizacao.sort_values('hours', ascending=False)
                    
                    # Visualizar tabela de utilização
                    st.subheader("Detalhamento por Recurso")
                    
                    # Preparar dataframe formatado para exibição
                    display_df = recursos_utilizacao.copy()
                    
                    # Converter horas para formato HH:MM
                    display_df['Horas'] = display_df['hours'].apply(format_hours_minutes)
                    
                    # Renomear e selecionar colunas para exibição
                    display_df = display_df.rename(columns={
                        'nome_completo': 'Colaborador',
                        'custo': 'Custo',
                        'perc_horas': '% Horas',
                        'perc_custo': '% Custo'
                    })
                    
                    # Formatar colunas
                    st.dataframe(
                        display_df[['Colaborador', 'Horas', 'Custo', '% Horas', '% Custo']].style.format({
                            'Custo': '€{:.2f}',
                            '% Horas': '{:.1f}%',
                            '% Custo': '{:.1f}%'
                        }).hide(axis='index')
                    )
                    
                    # Adicionar linha de totais
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total de Horas", format_hours_minutes(total_horas))
                    with col2:
                        st.metric("Custo Total", f"€{total_custo:.2f}")
                    
                    # Gráfico de distribuição de horas por recurso
                    st.subheader("Distribuição de Horas por Recurso")
                    
                    # Preparar dados com horas formatadas para exibição nos gráficos
                    recursos_utilizacao['hours_fmt'] = recursos_utilizacao['hours'].apply(format_hours_minutes)
                    
                    fig_horas = px.bar(
                        recursos_utilizacao,
                        x='nome_completo',
                        y='hours',
                        text=recursos_utilizacao['hours_fmt'],
                        title='Horas por Recurso',
                        color='perc_horas',
                        color_continuous_scale='Blues',
                        labels={
                            'nome_completo': 'Colaborador',
                            'hours': 'Horas Trabalhadas',
                            'perc_horas': '% do Total'
                        }
                    )
                    
                    fig_horas.update_layout(
                        xaxis_title='Colaborador',
                        yaxis_title='Horas Trabalhadas',
                        coloraxis_colorbar=dict(title='% do Total')
                    )
                    
                    st.plotly_chart(fig_horas, use_container_width=True)
                    
                    # Gráfico de distribuição de custos por recurso
                    st.subheader("Distribuição de Custos por Recurso")
                    
                    fig_custos = px.bar(
                        recursos_utilizacao,
                        x='nome_completo',
                        y='custo',
                        text=recursos_utilizacao['custo'].apply(lambda x: f"€{x:.0f}"),
                        title='Custo por Recurso',
                        color='perc_custo',
                        color_continuous_scale='Greens',
                        labels={
                            'nome_completo': 'Colaborador',
                            'custo': 'Custo (€)',
                            'perc_custo': '% do Total'
                        }
                    )
                    
                    fig_custos.update_layout(
                        xaxis_title='Colaborador',
                        yaxis_title='Custo (€)',
                        coloraxis_colorbar=dict(title='% do Total')
                    )
                    
                    st.plotly_chart(fig_custos, use_container_width=True)
                    
                else:
                    st.info("Não há registros de horas para este projeto ainda.")

                
                # Adicionar após a linha: st.markdown("---")
# Logo após a seção de fases do projeto (aproximadamente linha 814)

                st.markdown("---")
                st.subheader("📊 Consumo por Atividades e Categorias")

                # Verificar se existem entradas de timesheet para analisar
                if not entries.empty:
                    try:
                        # Carregar tabelas de categorias e atividades
                        categories_df = db_manager.query_to_df("SELECT * FROM task_categories")
                        activities_df = db_manager.query_to_df("SELECT * FROM activities")
                        
                        # Adicionar informações de categoria e atividade às entradas de timesheet
                        timesheet_completo = entries.copy()
                        
                        # Juntar com categorias (se category_id existir nas entradas)
                        if 'category_id' in timesheet_completo.columns:
                            timesheet_completo = timesheet_completo.merge(
                                categories_df[['task_category_id', 'task_category']],
                                left_on='category_id',
                                right_on='task_category_id',
                                how='left'
                            )
                        
                        # Juntar com atividades (se task_id existir nas entradas)
                        if 'task_id' in timesheet_completo.columns:
                            timesheet_completo = timesheet_completo.merge(
                                activities_df[['activity_id', 'activity_name']],
                                left_on='task_id',
                                right_on='activity_id',
                                how='left'
                            )
                        
                        # Juntar com informações de usuários para obter rates
                        timesheet_completo = timesheet_completo.merge(
                            users_df[['user_id', 'First_Name', 'Last_Name', 'rate_id']],
                            on='user_id',
                            how='left'
                        )
                        
                        # Função para calcular custo baseado na rate do usuário e horas extras
                        def calcular_custo(row):
                            if pd.isna(row['rate_id']):
                                return 0
                            
                            rate_info = rates_df[rates_df['rate_id'] == row['rate_id']]
                            if rate_info.empty:
                                return 0
                            
                            rate_value = float(rate_info['rate_cost'].iloc[0])
                            is_overtime = row.get('overtime', False)
                            
                            # Converter is_overtime para booleano
                            if isinstance(is_overtime, (int, float)):
                                is_overtime = bool(is_overtime)
                            elif isinstance(is_overtime, str):
                                is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                            
                            # Multiplicar por 2 se for hora extra
                            multiplicador = 2 if is_overtime else 1
                            return float(row['hours']) * rate_value * multiplicador
                        
                        timesheet_completo['custo_calculado'] = timesheet_completo.apply(calcular_custo, axis=1)
                        
                        # Criar abas para análise por categoria e por atividade
                        cat_act_tab1, cat_act_tab2 = st.tabs(["Por Categoria", "Por Atividade"])
                        
                        with cat_act_tab1:
                            # Verificar se há informações de categoria
                            if 'task_category' in timesheet_completo.columns:
                                # Agrupar por categoria
                                horas_por_categoria = timesheet_completo.groupby('task_category').agg({
                                    'hours': 'sum',
                                    'custo_calculado': 'sum'
                                }).reset_index()
                                
                                # Calcular percentuais
                                total_horas_cat = horas_por_categoria['hours'].sum()
                                total_custo_cat = horas_por_categoria['custo_calculado'].sum()
                                
                                horas_por_categoria['percentual_horas'] = horas_por_categoria['hours'].apply(
                                    lambda x: (x / total_horas_cat * 100) if total_horas_cat > 0 else 0
                                )
                                
                                horas_por_categoria['percentual_custo'] = horas_por_categoria['custo_calculado'].apply(
                                    lambda x: (x / total_custo_cat * 100) if total_custo_cat > 0 else 0
                                )
                                
                                # Substituir valores NaN na categoria
                                horas_por_categoria['task_category'] = horas_por_categoria['task_category'].fillna('Sem categoria')
                                
                                # Ordenar pelo número de horas (decrescente)
                                horas_por_categoria = horas_por_categoria.sort_values('hours', ascending=False)
                                
                                # Gráfico de horas por categoria
                                fig_cat = px.bar(
                                    horas_por_categoria,
                                    x='task_category',
                                    y='hours',
                                    text=horas_por_categoria['hours'].apply(format_hours_minutes),
                                    title='Horas por Categoria',
                                    color='percentual_horas',
                                    color_continuous_scale='Blues',
                                    labels={
                                        'task_category': 'Categoria',
                                        'hours': 'Horas',
                                        'percentual_horas': '% do Total'
                                    }
                                )
                                
                                fig_cat.update_layout(
                                    xaxis_title='Categoria',
                                    yaxis_title='Horas',
                                    xaxis={'categoryorder': 'total descending'}
                                )
                                
                                st.plotly_chart(fig_cat, use_container_width=True)
                                
                                # Tabela detalhada
                                display_cat_df = horas_por_categoria.copy()
                                display_cat_df['Horas'] = display_cat_df['hours'].apply(format_hours_minutes)
                                display_cat_df['Custo'] = display_cat_df['custo_calculado'].apply(lambda x: f"€{x:.2f}")
                                
                                st.dataframe(
                                    display_cat_df.rename(columns={
                                        'task_category': 'Categoria',
                                        'percentual_horas': '% Horas',
                                        'percentual_custo': '% Custo'
                                    })[['Categoria', 'Horas', '% Horas', 'Custo', '% Custo']].style.format({
                                        '% Horas': '{:.1f}%',
                                        '% Custo': '{:.1f}%'
                                    }).hide(axis='index')
                                )
                            else:
                                st.info("Não há informações de categoria disponíveis para este projeto.")
                        
                        with cat_act_tab2:
                            # Verificar se há informações de atividade
                            if 'activity_name' in timesheet_completo.columns:
                                # Agrupar por atividade
                                horas_por_atividade = timesheet_completo.groupby('activity_name').agg({
                                    'hours': 'sum',
                                    'custo_calculado': 'sum'
                                }).reset_index()
                                
                                # Calcular percentuais
                                total_horas_act = horas_por_atividade['hours'].sum()
                                total_custo_act = horas_por_atividade['custo_calculado'].sum()
                                
                                horas_por_atividade['percentual_horas'] = horas_por_atividade['hours'].apply(
                                    lambda x: (x / total_horas_act * 100) if total_horas_act > 0 else 0
                                )
                                
                                horas_por_atividade['percentual_custo'] = horas_por_atividade['custo_calculado'].apply(
                                    lambda x: (x / total_custo_act * 100) if total_custo_act > 0 else 0
                                )
                                
                                # Substituir valores NaN na atividade
                                horas_por_atividade['activity_name'] = horas_por_atividade['activity_name'].fillna('Sem atividade')
                                
                                # Ordenar pelo número de horas (decrescente)
                                horas_por_atividade = horas_por_atividade.sort_values('hours', ascending=False)
                                
                                # Gráfico de horas por atividade
                                # Gráfico de horas por atividade
                                fig_act = px.bar(
                                    horas_por_atividade,
                                    x='activity_name',
                                    y='hours',
                                    text=horas_por_atividade['hours'].apply(format_hours_minutes),
                                    title='Horas por Atividade',
                                    color='percentual_horas',
                                    color_continuous_scale='Greens',  # Usando uma escala diferente para diferenciar das categorias
                                    labels={
                                        'activity_name': 'Atividade',
                                        'hours': 'Horas',
                                        'percentual_horas': '% do Total'
                                    }
                                )

                                fig_act.update_layout(
                                    xaxis_title='Atividade',
                                    yaxis_title='Horas',
                                    xaxis={'categoryorder': 'total descending'}
                                )

                                st.plotly_chart(fig_act, use_container_width=True)
                                
                                # Tabela detalhada
                                display_act_df = horas_por_atividade.copy()
                                display_act_df['Horas'] = display_act_df['hours'].apply(format_hours_minutes)
                                display_act_df['Custo'] = display_act_df['custo_calculado'].apply(lambda x: f"€{x:.2f}")
                                
                                st.dataframe(
                                    display_act_df.rename(columns={
                                        'activity_name': 'Atividade',
                                        'percentual_horas': '% Horas',
                                        'percentual_custo': '% Custo'
                                    })[['Atividade', 'Horas', '% Horas', 'Custo', '% Custo']].style.format({
                                        '% Horas': '{:.1f}%',
                                        '% Custo': '{:.1f}%'
                                    }).hide(axis='index')
                                )
                                
                                # Adicionar visualização combinada (categoria + atividade)
                                # Adicionar visualização combinada (categoria + atividade)

                            else:
                                st.info("Não há informações de atividade disponíveis para este projeto.")
                    except Exception as e:
                        st.warning(f"Não foi possível gerar a análise por categorias e atividades: {str(e)}")
                        # Se quiser ver o traceback completo, descomente a linha abaixo
                        # st.sidebar.error(traceback.format_exc())
                else:
                    st.info("Não há registros de horas para analisar neste projeto.")
                
                st.markdown("---")
                st.subheader("📅 Fases do Projeto")

                # Obter informações das fases
                phases_info = integrate_phases_with_project_reports(projeto_id)

                if phases_info and phases_info['total_phases'] > 0:
                    # Mostrar resumo das fases
                    st.write(f"Este projeto possui {phases_info['total_phases']} fases definidas.")
                    
                    # Criar dados para a tabela de fases
                    phases_data = []
                    for phase in phases_info['phases']:
                        phases_data.append({
                            'Nome': phase['name'],
                            'Início': phase['start_date'].strftime('%d/%m/%Y'),
                            'Término': phase['end_date'].strftime('%d/%m/%Y'),
                            'Horas': format_hours_minutes(phase['total_hours']),
                            'Custo': f"€{phase['total_cost']:.2f}",
                            'Status': phase['status'].capitalize()
                        })
                    
                    # Exibir tabela de fases
                    phases_df = pd.DataFrame(phases_data)
                    st.dataframe(phases_df)
                    
                    # Criar timeline das fases
                    if len(phases_info['phases']) > 0:
                        st.subheader("Timeline das Fases")
                        
                        # Preparar dados para o gráfico de Gantt
                        gantt_data = []
                        for phase in phases_info['phases']:
                            gantt_data.append({
                                'Fase': phase['name'],
                                'Início': phase['start_date'],
                                'Fim': phase['end_date'],
                                'Horas': phase['total_hours'],
                                'Custo': phase['total_cost']
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
                            x0=pd.to_datetime(projeto_info['start_date']),
                            y0=-0.5,
                            x1=pd.to_datetime(projeto_info['start_date']),
                            y1=len(phases_info['phases']) - 0.5,
                            line=dict(color='red', width=2, dash='dash'),
                            name='Início do Projeto'
                        )
                        
                        fig.add_shape(
                            type='line',
                            x0=pd.to_datetime(projeto_info['end_date']),
                            y0=-0.5,
                            x1=pd.to_datetime(projeto_info['end_date']),
                            y1=len(phases_info['phases']) - 0.5,
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
                        
                        # Adicionar análise de progresso das fases
                        st.subheader("Progresso das Fases")
                        
                        # Verificar progresso de cada fase com base na data atual
                        today = datetime.now().date()
                        progress_data = []
                        
                        for phase in phases_info['phases']:
                            phase_start = phase['start_date'].date()
                            phase_end = phase['end_date'].date()
                            phase_duration = (phase_end - phase_start).days
                            
                            # Calcular progresso esperado
                            if today < phase_start:
                                expected_progress = 0  # Fase não iniciada
                            elif today > phase_end:
                                expected_progress = 100  # Fase concluída
                            else:
                                days_passed = (today - phase_start).days
                                expected_progress = (days_passed / phase_duration * 100) if phase_duration > 0 else 0
                            
                            # Determinar status da fase
                            if phase['status'] == 'completed':
                                actual_progress = 100
                                status = "Concluída"
                            elif phase['status'] == 'active':
                                # Para fases ativas, o progresso é proporcional ao tempo
                                actual_progress = expected_progress
                                status = "Em Andamento"
                            else:
                                actual_progress = 0
                                status = "Pendente"
                            
                            # Determinar se está no prazo
                            on_schedule = "No Prazo" if actual_progress >= expected_progress else "Atrasada"
                            
                            progress_data.append({
                                'Fase': phase['name'],
                                'Progresso Esperado (%)': expected_progress,
                                'Progresso Atual (%)': actual_progress,
                                'Status': status,
                                'Situação': on_schedule
                            })
                        
                        # Exibir tabela de progresso
                        progress_df = pd.DataFrame(progress_data)
                        st.dataframe(progress_df)
                        
                        # Exibir gráfico de progresso das fases
                        fig_progress = go.Figure()
                        
                        # Adicionar barras de progresso esperado
                        fig_progress.add_trace(go.Bar(
                            y=progress_df['Fase'],
                            x=progress_df['Progresso Esperado (%)'],
                            name='Progresso Esperado',
                            orientation='h',
                            marker=dict(color='rgba(55, 83, 109, 0.7)'),
                            hovertemplate='%{x:.1f}%<extra>Esperado</extra>'
                        ))
                        
                        # Adicionar barras de progresso atual
                        fig_progress.add_trace(go.Bar(
                            y=progress_df['Fase'],
                            x=progress_df['Progresso Atual (%)'],
                            name='Progresso Atual',
                            orientation='h',
                            marker=dict(color='rgba(26, 118, 255, 0.7)'),
                            hovertemplate='%{x:.1f}%<extra>Atual</extra>'
                        ))
                        
                        # Ajustar layout
                        fig_progress.update_layout(
                            title='Comparação de Progresso das Fases',
                            xaxis_title='Progresso (%)',
                            yaxis_title='Fase',
                            barmode='group',
                            height=400,
                            xaxis=dict(range=[0, 100])
                        )
                        
                        st.plotly_chart(fig_progress, use_container_width=True)
                
                        
                else:
                    st.info("Este projeto não possui fases definidas. Você pode adicionar fases através do menu 'Relatórios > Fases de Projetos'.")
        # Adicionar botão de exportação de relatório PDF
        add_report_export_button(projeto_info, client, db_manager)       
    # Tratamento global de exceções para não deixar a tela em branco
    except Exception as e:
        st.error("Ocorreu um erro ao carregar o relatório de projetos:")
        st.error(str(e))
        st.error(traceback.format_exc())
        
        # Mostrar instruções de recuperação
        st.warning("Tente as seguintes ações para resolver o problema:")
        st.write("1. Verifique se você tem acesso aos dados necessários")
        st.write("2. Certifique-se de que está atribuído a pelo menos uma equipe no sistema")
        st.write("3. Contate o administrador do sistema se o problema persistir")

if __name__ == "__main__":
    reports_page()