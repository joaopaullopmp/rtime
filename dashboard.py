# dashboard.py
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime, timedelta
import calendar
from database_manager import DatabaseManager, UserManager, ProjectManager, ClientManager, GroupManager
from annual_targets import AnnualTargetManager
from collaborator_targets import CollaboratorTargetCalculator
from billing_manager import BillingManager

# Exportar a função dashboard_page para ser acessada de outros módulos
__all__ = ['dashboard_page']

def dashboard_page():
    """
    Dashboard principal com indicadores de performance da empresa
    """
    st.title("Dashboard de Indicadores")
    
    try:
        # Inicializar gerenciadores
        db_manager = DatabaseManager()
        annual_target_manager = AnnualTargetManager()
        collaborator_target_calculator = CollaboratorTargetCalculator()
        billing_manager = BillingManager()
        
        # Selecionar mês e ano atuais para referência
        now = datetime.now()
        current_month = now.month
        current_year = now.year
        
        # Barra lateral com configurações gerais do dashboard
        with st.sidebar:
            st.header("Filtros Gerais")
            
            # Seleção de período
            col1, col2 = st.columns(2)
            with col1:
                month = st.selectbox(
                    "Mês",
                    range(1, 13),
                    format_func=lambda x: calendar.month_name[x],
                    index=current_month - 1,
                    key="dash_month"
                )
            with col2:
                year = st.selectbox(
                    "Ano",
                    range(current_year - 1, current_year + 3),
                    index=1,
                    key="dash_year"
                )
            
            # Carregar dados das equipes
            groups_df = db_manager.query_to_df("SELECT * FROM groups WHERE active = 1")
            
            # Filtro de equipe
            teams = ["Todas"] + groups_df["group_name"].tolist()
            selected_team = st.selectbox(
                "Equipa",
                options=teams,
                index=0,
                key="dash_team"
            )
        
        # Tabs do dashboard
        tab1, tab2, tab3 = st.tabs([
            "Indicadores de Colaboradores", 
            "Indicadores de Projetos", 
            "Indicadores de Faturação"
        ])
        
        with tab1:
            show_collaborator_indicators(
                db_manager, 
                collaborator_target_calculator, 
                month, 
                year, 
                selected_team
            )
        
        with tab2:
            show_project_indicators(
                db_manager, 
                annual_target_manager, 
                month, 
                year, 
                selected_team
            )
        
        with tab3:
            show_revenue_indicators(
                db_manager, 
                annual_target_manager,
                billing_manager,
                month, 
                year, 
                selected_team
            )
    
    except Exception as e:
        st.error(f"Erro ao carregar o dashboard: {str(e)}")
        import traceback
        st.code(traceback.format_exc(), language="python")

def show_collaborator_indicators(db_manager, collaborator_target_calculator, month, year, selected_team):
    """
    Mostra indicadores de performance dos colaboradores
    com metas de 87.5% para ocupação e 75% para horas faturáveis
    considerando os feriados de Portugal
    """
    # Exibir título com informações de referência (equipe e mês)
    month_name = calendar.month_name[month]
    if selected_team == "Todas":
        st.subheader(f"Indicadores de Colaboradores - Todas as Equipas - {month_name}/{year}")
    
    # Carregar dados necessários
    timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
    users_df = db_manager.query_to_df("SELECT * FROM utilizadores WHERE active = 1")
    groups_df = db_manager.query_to_df("SELECT * FROM groups")
    
    # Filtrar por equipe se necessário
    if selected_team != "Todas":
        # Exibir título com equipe específica
        st.subheader(f"Indicadores de Colaboradores - Equipa: {selected_team} - {month_name}/{year}")
        
        # Encontrar os usuários que pertencem à equipa selecionada
        filtered_users = []
        for _, user in users_df.iterrows():
            try:
                if isinstance(user['groups'], str):
                    user_groups = eval(user['groups'])
                else:
                    user_groups = user['groups'] if user['groups'] is not None else []
                
                # Converter para lista se for outro tipo
                if not isinstance(user_groups, list):
                    if isinstance(user_groups, dict):
                        user_groups = list(user_groups.values())
                    else:
                        user_groups = [user_groups]
                
                # Verificar se o usuário pertence à equipe selecionada
                if selected_team in user_groups:
                    filtered_users.append(user['user_id'])
            except Exception as e:
                st.warning(f"Erro ao processar grupos do usuário {user['First_Name']} {user['Last_Name']}: {str(e)}")
        
        # Filtrar usuários
        users_df = users_df[users_df['user_id'].isin(filtered_users)]
    
    if users_df.empty:
        st.warning("Não foram encontrados colaboradores com os filtros selecionados.")
        return
    
    # Calcular dias úteis do mês considerando feriados
    from report_utils import get_feriados_portugal
    
    primeiro_dia = datetime(year, month, 1)
    ultimo_dia = datetime(year, month, calendar.monthrange(year, month)[1], 23, 59, 59)
    
    # Obter lista de feriados para o ano
    feriados = get_feriados_portugal(year)
    
    # Identificar feriados no mês atual
    feriados_no_mes = []
    for feriado in feriados:
        if feriado.month == month and feriado.year == year:
            feriados_no_mes.append(feriado)
    
    # Exibir informações sobre feriados do mês
    if feriados_no_mes:
        feriados_info = ", ".join([f"{f.day}/{f.month}" for f in feriados_no_mes])
        st.info(f"⚠️ Feriados considerados em {month_name}/{year}: {feriados_info}")
    
    # Calcular dias úteis (de trabalho) no mês, excluindo feriados
    dias_uteis = 0
    data_atual = primeiro_dia
    while data_atual <= ultimo_dia:
        # Se não for sábado (5) nem domingo (6) e não for feriado
        if data_atual.weekday() < 5 and data_atual.date() not in feriados:
            dias_uteis += 1
        data_atual += timedelta(days=1)
    
    # Horas úteis totais (considerando 8 horas por dia útil)
    horas_uteis_mes = dias_uteis * 8
    
    # Para cada colaborador, calcular indicadores
    collaborator_indicators = []
    
    for _, user in users_df.iterrows():
        # Obter dados de metas do colaborador
        user_targets = collaborator_target_calculator.get_user_targets(user['user_id'], year)
        
        # Filtrar entradas de timesheet para o colaborador e mês específico
        try:
            user_timesheet = timesheet_df[
                (timesheet_df['user_id'] == user['user_id']) & 
                (pd.to_datetime(timesheet_df['start_date'], format='mixed').dt.month == month) &
                (pd.to_datetime(timesheet_df['start_date'], format='mixed').dt.year == year)
            ]
        except Exception as e:
            # Fallback em caso de erro de conversão de data
            st.warning(f"Erro ao filtrar timesheet para {user['First_Name']} {user['Last_Name']}: {e}")
            user_timesheet = pd.DataFrame()
        
        # Calcular horas realizadas (total de horas registradas)
        total_hours = user_timesheet['hours'].sum() if not user_timesheet.empty else 0
        
        # Calcular horas faturáveis (apenas entradas marcadas como billable=True)
        billable_hours = user_timesheet[user_timesheet['billable'] == True]['hours'].sum() if not user_timesheet.empty else 0
        
        # Calcular percentuais de ocupação e faturabilidade baseados nas horas realizadas
        # Ocupação = (Horas realizadas / Horas úteis do mês) * 100
        occupation_percentage = (total_hours / horas_uteis_mes * 100) if horas_uteis_mes > 0 else 0
        
        # Faturabilidade = (Horas faturáveis / Horas úteis do mês) * 100
        billable_percentage = (billable_hours / horas_uteis_mes * 100) if horas_uteis_mes > 0 else 0
        
        # Metas definidas
        target_occupation = 87.5  # Meta de ocupação (87,5%)
        target_billable = 75.0    # Meta de horas faturáveis (75.0%)
        
        # Determinar cores dos indicadores - simplificado para verde/vermelho
        occupation_color = "green" if occupation_percentage >= target_occupation else "red"
        billable_color = "green" if billable_percentage >= target_billable else "red"
        
        # Adicionar ao array de indicadores
        collaborator_indicators.append({
            "user_id": user['user_id'],
            "name": f"{user['First_Name']} {user['Last_Name']}",
            "occupation_percentage": occupation_percentage,
            "billable_percentage": billable_percentage,
            "occupation_color": occupation_color,
            "billable_color": billable_color,
            "target_occupation": target_occupation,
            "target_billable": target_billable,
            "total_hours": total_hours,  # Quantidade de horas realizadas
            "billable_hours": billable_hours  # Quantidade de horas faturáveis
        })
    
    # Converter para DataFrame
    indicators_df = pd.DataFrame(collaborator_indicators)
    
    # Layout de indicadores
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Ocupação")
        
        # Gráfico de ocupação com cores
        fig_occupation = px.bar(
            indicators_df.sort_values("occupation_percentage", ascending=False),
            x="name",
            y="occupation_percentage",
            title="Percentual de Ocupação por Colaborador",
            labels={"name": "Colaborador", "occupation_percentage": "% Ocupação"},
            color_discrete_map={"green": "#4CAF50", "red": "#F44336"},
            color="occupation_color",
            text_auto='.1f'  # Adicionar rótulos automáticos com uma casa decimal
        )
        
        # Adicionar linha de referência para meta com anotação
        fig_occupation.add_shape(
            type="line",
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            y0=target_occupation,
            y1=target_occupation,
            line=dict(color="black", width=2, dash="dash"),
        )
        
        # Adicionar anotação para a linha de meta
        fig_occupation.add_annotation(
            x=len(indicators_df)-1,
            y=target_occupation,
            text=f"Meta: {target_occupation}%",
            showarrow=False,
            yshift=10,
            font=dict(
                family="Arial",
                size=12,
                color="black"
            ),
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="black",
            borderwidth=1,
            borderpad=4
        )
        
        # Ajustar configurações do layout
        fig_occupation.update_layout(
            xaxis_tickangle=-45,
            height=500,
            uniformtext_minsize=8,  # Tamanho mínimo do texto dos rótulos
            uniformtext_mode='hide'  # Esconder rótulos muito pequenos
        )
        
        # Ajustar posição dos rótulos
        fig_occupation.update_traces(textposition='outside')
        
        st.plotly_chart(fig_occupation, use_container_width=True)
    
    with col2:
        st.markdown("### Faturabilidade")
        
        # Gráfico de faturabilidade com cores
        fig_billable = px.bar(
            indicators_df.sort_values("billable_percentage", ascending=False),
            x="name",
            y="billable_percentage",
            title="Percentual de Horas Faturáveis por Colaborador",
            labels={"name": "Colaborador", "billable_percentage": "% Faturável"},
            color_discrete_map={"green": "#4CAF50", "red": "#F44336"},
            color="billable_color",
            text_auto='.1f'  # Adicionar rótulos automáticos com uma casa decimal
        )
        
        # Adicionar linha de referência para meta
        fig_billable.add_shape(
            type="line",
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            y0=target_billable,
            y1=target_billable,
            line=dict(color="black", width=2, dash="dash"),
        )
        
        # Adicionar anotação para a linha de meta
        fig_billable.add_annotation(
            x=len(indicators_df)-1,
            y=target_billable,
            text=f"Meta: {target_billable}%",
            showarrow=False,
            yshift=10,
            font=dict(
                family="Arial",
                size=12,
                color="black"
            ),
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="black",
            borderwidth=1,
            borderpad=4
        )
        
        # Ajustar configurações do layout
        fig_billable.update_layout(
            xaxis_tickangle=-45,
            height=500,
            uniformtext_minsize=8,  # Tamanho mínimo do texto dos rótulos
            uniformtext_mode='hide'  # Esconder rótulos muito pequenos
        )
        
        # Ajustar posição dos rótulos
        fig_billable.update_traces(textposition='outside')
        
        st.plotly_chart(fig_billable, use_container_width=True)
    
    # Tabela de indicadores
    st.markdown("### Indicadores Detalhados")
    
    # Criar colunas estilizadas para os indicadores
    display_df = indicators_df.copy()
    
    # Definir estilo CSS para tornar a tabela mais atraente
    table_style = """
    <style>
        .styled-table {
            width: 100%;
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 14px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
        }
        .styled-table thead tr {
            background-color: #2c3e50;
            color: white;
            text-align: left;
            font-weight: bold;
        }
        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
            text-align: left;
        }
        .styled-table th:first-child,
        .styled-table td:first-child {
            text-align: left;
        }
        .styled-table tbody tr {
            border-bottom: 1px solid #dddddd;
        }
        .styled-table tbody tr:nth-of-type(even) {
            background-color: #f3f3f3;
        }
        .styled-table tbody tr:last-of-type {
            border-bottom: 2px solid #2c3e50;
        }
        .indicator-dot {
            border-radius: 50%;
            width: 15px;
            height: 15px;
            display: inline-block;
            margin-right: 6px;
            vertical-align: middle;
        }
        .number-cell {
            font-family: 'Courier New', monospace;
            font-weight: bold;
        }
    </style>
    """
    
    # Reformatar a exibição das células com os indicadores
    display_df['Ocupação'] = display_df.apply(
        lambda x: f'<div><span class="indicator-dot" style="background-color: {"#4CAF50" if x.occupation_color == "green" else "#F44336"}"></span><span class="number-cell">{x.occupation_percentage:.1f}%</span></div>',
        axis=1
    )
    display_df['Faturável'] = display_df.apply(
        lambda x: f'<div><span class="indicator-dot" style="background-color: {"#4CAF50" if x.billable_color == "green" else "#F44336"}"></span><span class="number-cell">{x.billable_percentage:.1f}%</span></div>',
        axis=1
    )
    
    # Formatar células numéricas com classe especial
    display_df['Horas Realizadas'] = display_df.apply(
        lambda x: f'<span class="number-cell">{x.total_hours:.2f}</span>',
        axis=1
    )
    display_df['Horas Faturáveis'] = display_df.apply(
        lambda x: f'<span class="number-cell">{x.billable_hours:.2f}</span>',
        axis=1
    )
    
    # Mostrar a tabela com as novas colunas e estilos
    st.markdown(table_style, unsafe_allow_html=True)
    st.write(
        '<table class="styled-table">' +
        display_df[['name', 'Ocupação', 'Horas Realizadas', 'Faturável', 'Horas Faturáveis']]
        .rename(columns={'name': 'Colaborador'})
        .to_html(escape=False, index=False)
        .replace('<table', '<table class="styled-table"'),
        unsafe_allow_html=True
    )
    
    # Legenda de cores e informações sobre as metas
    st.markdown("### Informações e Legenda")
    
    # Criar um card informativo sobre o mês de referência
    st.markdown(f"""
    <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 15px;">
        <h4 style="margin-top: 0;">Resumo do Mês</h4>
        <p><strong>Mês de Referência:</strong> {month_name}/{year}</p>
        <p><strong>Dias Úteis:</strong> {dias_uteis} dias úteis</p>
        <p><strong>Horas Úteis:</strong> {horas_uteis_mes} horas úteis</p>
        <p><strong>Feriados:</strong> {len(feriados_no_mes)} dia(s) de feriado no mês</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("#### Metas dos Colaboradores")
    st.markdown(f"**Meta de Ocupação**: 87.5% das horas úteis totais do colaborador (baseado em {horas_uteis_mes} horas úteis no mês)")
    st.markdown(f"**Meta de Horas Faturáveis**: 75% das horas úteis totais do colaborador")
    
    # Esquema de cores melhorado
    st.markdown("""
    <div style="display: flex; margin-top: 15px;">
        <div style="flex: 1; padding: 10px; margin-right: 10px; background-color: rgba(76, 175, 80, 0.1); border-left: 4px solid #4CAF50; border-radius: 4px;">
            <div style="display: flex; align-items: center;">
                <div style="background-color: #4CAF50; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                <strong>Igual ou acima da meta</strong>
            </div>
        </div>
        <div style="flex: 1; padding: 10px; background-color: rgba(244, 67, 54, 0.1); border-left: 4px solid #F44336; border-radius: 4px;">
            <div style="display: flex; align-items: center;">
                <div style="background-color: #F44336; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                <strong>Abaixo da meta</strong>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import calendar
import io

def show_project_indicators(db_manager, annual_target_manager, month, year, selected_team):
    """
    Mostra indicadores de performance dos projetos com interface melhorada,
    incluindo rótulos nos gráficos e tabela formatada com dados detalhados
    """
    # Exibir título com informações de referência (equipe e mês)
    month_name = calendar.month_name[month]
    if selected_team == "Todas":
        st.subheader(f"Indicadores de Projetos - Todas as Equipas - {month_name}/{year}")
    else:
        st.subheader(f"Indicadores de Projetos - Equipa: {selected_team} - {month_name}/{year}")
    
    # Carregar dados necessários - carregando todos os projetos (ativos e inativos)
    projects_df = db_manager.query_to_df("SELECT * FROM projects")
    clients_df = db_manager.query_to_df("SELECT * FROM clients")
    timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
    users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
    rates_df = db_manager.query_to_df("SELECT * FROM rates")
    
    # Filtros adicionais para esta seção
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Filtro de status do projeto (ativo/inativo)
        status_options = ["Todos", "Ativos", "Inativos"]
        selected_status = st.selectbox(
            "Status do Projeto",
            options=status_options,
            index=1  # Padrão: mostrar apenas projetos ativos
        )
    
    with col2:
        # Filtro de tipo de projeto
        project_types = ["Todos"] + sorted(projects_df["project_type"].unique().tolist())
        selected_type = st.selectbox(
            "Tipo de Projeto",
            options=project_types,
            index=0
        )
    
    with col3:
        # Filtro de projeto específico
        project_options = ["Todos"] + sorted(projects_df["project_name"].tolist())
        selected_project = st.selectbox(
            "Projeto",
            options=project_options,
            index=0
        )
    
    # Aplicar filtros
    filtered_projects = projects_df.copy()
    
    # Filtrar por status (ativo/inativo)
    if selected_status == "Ativos":
        filtered_projects = filtered_projects[filtered_projects["status"].str.lower() == "active"]
    elif selected_status == "Inativos":
        filtered_projects = filtered_projects[filtered_projects["status"].str.lower() != "active"]
    
    # Filtrar por equipe
    if selected_team != "Todas":
        # Encontrar o ID do grupo
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        group_id = groups_df[groups_df["group_name"] == selected_team]["id"].iloc[0] if not groups_df.empty else None
        
        if group_id:
            filtered_projects = filtered_projects[filtered_projects["group_id"] == group_id]
    
    # Filtrar por tipo de projeto
    if selected_type != "Todos":
        filtered_projects = filtered_projects[filtered_projects["project_type"] == selected_type]
    
    # Filtrar por projeto específico
    if selected_project != "Todos":
        filtered_projects = filtered_projects[filtered_projects["project_name"] == selected_project]
    
    if filtered_projects.empty:
        st.warning("Não foram encontrados projetos com os filtros selecionados.")
        return
    
    # Calcular indicadores de projetos
    project_indicators = []
    
    # Função para calcular dias úteis entre duas datas
    def calcular_dias_uteis(data_inicio, data_fim):
        """Calcula os dias úteis entre duas datas"""
        dias_uteis = 0
        data_atual = data_inicio
        
        while data_atual <= data_fim:
            if data_atual.weekday() < 5:  # 0-4 são dias úteis (seg-sex)
                dias_uteis += 1
            data_atual += timedelta(days=1)
        
        return dias_uteis
    
    # Função para calcular a proporção de dias úteis dentro de um mês
    def calcular_proporcao_mes(start_date, end_date, mes_ref, ano_ref):
        """Calcula a proporção de dias úteis do projeto no mês de referência"""
        # Se o projeto começa após ou termina antes do mês de referência, retorna 0
        inicio_mes_ref = datetime(ano_ref, mes_ref, 1)
        ultimo_dia_mes_ref = calendar.monthrange(ano_ref, mes_ref)[1]
        fim_mes_ref = datetime(ano_ref, mes_ref, ultimo_dia_mes_ref, 23, 59, 59)
        
        if start_date > fim_mes_ref or end_date < inicio_mes_ref:
            return 0
        
        # Ajustar datas para o período do mês
        periodo_inicio = max(start_date, inicio_mes_ref)
        periodo_fim = min(end_date, fim_mes_ref)
        
        # Calcular dias úteis no período do mês
        dias_uteis_periodo = calcular_dias_uteis(periodo_inicio, periodo_fim)
        
        # Calcular dias úteis totais no projeto
        dias_uteis_projeto = calcular_dias_uteis(start_date, end_date)
        
        if dias_uteis_projeto == 0:
            return 0
        
        return dias_uteis_periodo / dias_uteis_projeto
    
    # Definir o início e fim do mês de referência
    inicio_mes = datetime(year, month, 1)
    ultimo_dia = calendar.monthrange(year, month)[1]
    fim_mes = datetime(year, month, ultimo_dia, 23, 59, 59)
    
    # Início do ano
    inicio_ano = datetime(year, 1, 1)
    fim_ano = datetime(year, 12, 31, 23, 59, 59)
    
    # Data atual para considerar apenas períodos já passados na análise anual
    data_atual = datetime.now()
    
    # Converter datas em timesheet para datetime para facilitar filtros posteriores
    if 'start_date' in timesheet_df.columns:
        try:
            timesheet_df['start_date_dt'] = pd.to_datetime(timesheet_df['start_date'], format='mixed')
        except Exception as e:
            st.warning(f"Erro ao converter datas: {e}")
            timesheet_df['start_date_dt'] = pd.to_datetime('2000-01-01')
    
    for _, project in filtered_projects.iterrows():
        # Converter datas de início e fim para datetime
        try:
            start_date = pd.to_datetime(project['start_date'])
            end_date = pd.to_datetime(project['end_date'])
        except:
            continue  # Pular projetos com datas inválidas
        
        # Obter registros de timesheet para o projeto
        project_entries = timesheet_df[timesheet_df["project_id"] == project["project_id"]]
        
        # Verificar se existem dados migrados (horas_realizadas_mig e custo_realizado_mig)
        horas_migradas = 0
        if 'horas_realizadas_mig' in project and not pd.isna(project['horas_realizadas_mig']):
            horas_migradas = float(project['horas_realizadas_mig'])
        
        custo_migrado = 0
        if 'custo_realizado_mig' in project and not pd.isna(project['custo_realizado_mig']):
            custo_migrado = float(project['custo_realizado_mig'])
        
        # Calcular dias úteis totais do projeto
        dias_uteis_projeto = calcular_dias_uteis(start_date, end_date)
        
        # Calcular a proporção de dias úteis no mês de referência em relação ao total do projeto
        proporcao_mes = calcular_proporcao_mes(start_date, end_date, month, year)
        
        # CÁLCULO PARA O MÊS ATUAL (apenas mês selecionado)
        # Filtrar entradas do mês atual
        try:
            month_entries = project_entries[
                (pd.to_datetime(project_entries["start_date"], format='mixed') >= inicio_mes) &
                (pd.to_datetime(project_entries["start_date"], format='mixed') <= fim_mes)
            ]
        except Exception as e:
            st.warning(f"Erro ao filtrar entradas do mês: {e}")
            month_entries = pd.DataFrame()
        
        # Calcular horas regulares e extras para o mês
        month_hours_regular = 0
        month_hours_extra = 0
        
        if not month_entries.empty:
            month_hours_regular = month_entries[~month_entries['overtime'].astype(bool)]['hours'].sum()
            # Para horas extras, calcular o dobro para contabilização
            month_hours_extra_original = month_entries[month_entries['overtime'].astype(bool)]['hours'].sum()
            month_hours_extra = month_hours_extra_original * 2
        
        # Total de horas no mês (regulares + extras*2)
        month_hours = float(month_hours_regular + month_hours_extra)
        
        # Calcular custo realizado para o mês
        month_cost = 0
        if not month_entries.empty:
            for _, entry in month_entries.iterrows():
                try:
                    user_id = entry['user_id']
                    hours = float(entry['hours'])
                    is_overtime = entry.get('overtime', False)
                    
                    # Converter para booleano
                    if isinstance(is_overtime, (int, float)):
                        is_overtime = bool(is_overtime)
                    elif isinstance(is_overtime, str):
                        is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                    
                    # Obter rate para o usuário
                    rate_value = None
                    if 'rate_value' in entry and not pd.isna(entry['rate_value']):
                        rate_value = float(entry['rate_value'])
                    else:
                        user_info = users_df[users_df['user_id'] == user_id]
                        
                        if not user_info.empty and not pd.isna(user_info['rate_id'].iloc[0]):
                            rate_id = user_info['rate_id'].iloc[0]
                            rate_info = rates_df[rates_df['rate_id'] == rate_id]
                            
                            if not rate_info.empty:
                                rate_value = float(rate_info['rate_cost'].iloc[0])
                    
                    # Calcular custo com base no rate obtido
                    if rate_value:
                        entry_cost = hours * rate_value
                        
                        # Se for hora extra, multiplicar por 2
                        if is_overtime:
                            month_cost += entry_cost * 2  # Dobro para horas extras
                        else:
                            month_cost += entry_cost  # Normal para horas regulares
                except Exception as e:
                    pass
        
        # Meta mensal proporcional à duração do projeto
        month_budget_hours = float(project["total_hours"]) * proporcao_mes if pd.notna(project["total_hours"]) else 0
        month_budget_cost = float(project["total_cost"]) * proporcao_mes if pd.notna(project["total_cost"]) else 0
        
        # Calcular percentual de conclusão mensal
        month_percentage = (month_hours / month_budget_hours * 100) if month_budget_hours > 0 else 0
        
        # CÁLCULO PARA O ANO (acumulado desde o início do ano até agora)
        # Filtrar entradas do ano até a data atual
        try:
            year_entries = project_entries[
                (pd.to_datetime(project_entries["start_date"], format='mixed') >= inicio_ano) &
                (pd.to_datetime(project_entries["start_date"], format='mixed') <= data_atual)
            ]
        except Exception as e:
            st.warning(f"Erro ao filtrar entradas do ano: {e}")
            year_entries = pd.DataFrame()
        
        # Calcular horas regulares e extras para o ano
        year_hours_regular = 0
        year_hours_extra = 0
        
        if not year_entries.empty:
            year_hours_regular = year_entries[~year_entries['overtime'].astype(bool)]['hours'].sum()
            # Para horas extras, calcular o dobro para contabilização
            year_hours_extra_original = year_entries[year_entries['overtime'].astype(bool)]['hours'].sum()
            year_hours_extra = year_hours_extra_original * 2
        
        # Total de horas no ano (regulares + extras*2 + horas migradas)
        year_hours = float(year_hours_regular + year_hours_extra + horas_migradas)
        
        # Calcular custo realizado para o ano
        year_cost = 0
        if not year_entries.empty:
            for _, entry in year_entries.iterrows():
                try:
                    user_id = entry['user_id']
                    hours = float(entry['hours'])
                    is_overtime = entry.get('overtime', False)
                    
                    # Converter para booleano
                    if isinstance(is_overtime, (int, float)):
                        is_overtime = bool(is_overtime)
                    elif isinstance(is_overtime, str):
                        is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                    
                    # Obter rate para o usuário
                    rate_value = None
                    if 'rate_value' in entry and not pd.isna(entry['rate_value']):
                        rate_value = float(entry['rate_value'])
                    else:
                        user_info = users_df[users_df['user_id'] == user_id]
                        
                        if not user_info.empty and not pd.isna(user_info['rate_id'].iloc[0]):
                            rate_id = user_info['rate_id'].iloc[0]
                            rate_info = rates_df[rates_df['rate_id'] == rate_id]
                            
                            if not rate_info.empty:
                                rate_value = float(rate_info['rate_cost'].iloc[0])
                    
                    # Calcular custo com base no rate obtido
                    if rate_value:
                        entry_cost = hours * rate_value
                        
                        # Se for hora extra, multiplicar por 2
                        if is_overtime:
                            year_cost += entry_cost * 2  # Dobro para horas extras
                        else:
                            year_cost += entry_cost  # Normal para horas regulares
                except Exception as e:
                    pass
        
        # Adicionar custo migrado ao total do ano
        year_cost += custo_migrado
        
        # Calcular a proporção de dias úteis entre o início do projeto e hoje
        # em relação ao total de dias úteis do projeto
        if start_date <= data_atual:
            dias_uteis_ate_hoje = calcular_dias_uteis(start_date, min(data_atual, end_date))
            proporcao_atual = dias_uteis_ate_hoje / dias_uteis_projeto if dias_uteis_projeto > 0 else 0
        else:
            proporcao_atual = 0
            
        # Orçamento total de horas e custo do projeto
        year_budget_hours = float(project["total_hours"]) if pd.notna(project["total_hours"]) else 0.0
        year_budget_cost = float(project["total_cost"]) if pd.notna(project["total_cost"]) else 0.0
        
        # Calcular percentual de conclusão anual do projeto (incluindo horas migradas)
        year_percentage = (year_hours / year_budget_hours * 100) if year_budget_hours > 0 else 0
        
        # Calcular percentual esperado de conclusão, baseado no tempo decorrido
        expected_percentage = proporcao_atual * 100
        
        # Comparar o percentual executado com o esperado
        # Se percentage_ratio > 1, o projeto está consumindo horas mais rápido que o planejado
        if expected_percentage > 0:
            percentage_ratio = year_percentage / expected_percentage
        else:
            percentage_ratio = 0
        
        # Determinar cores dos indicadores utilizando as métricas originais
        # Meta mensal - apenas dados do mês atual
        if month_percentage > 80:
            month_color = "red"
        elif month_percentage >= 60:
            month_color = "yellow"
        else:
            month_color = "green"
            
        # Meta anual - dados acumulados do ano, incluindo migrados
        if year_percentage > 60:
            year_color = "red"
        elif year_percentage >= 40:
            year_color = "yellow"
        else:
            year_color = "green"
        
        # Obter nome do cliente
        client_name = clients_df[clients_df["client_id"] == project["client_id"]]["name"].iloc[0] if not clients_df.empty else "Cliente Desconhecido"
        
        # Converter datas para formato datetime e aplicar formatação
        try:
            start_date_fmt = pd.to_datetime(project["start_date"]).strftime('%d/%m/%Y') if pd.notna(project["start_date"]) else "N/A"
            end_date_fmt = pd.to_datetime(project["end_date"]).strftime('%d/%m/%Y') if pd.notna(project["end_date"]) else "N/A"
        except:
            start_date_fmt = "N/A"
            end_date_fmt = "N/A"
        
        # Adicionar ao array de indicadores
        project_indicators.append({
            "project_id": project["project_id"],
            "project_name": project["project_name"],
            "client_name": client_name,
            "project_type": project["project_type"],
            "status": project["status"],
            "start_date": start_date_fmt,
            "end_date": end_date_fmt,
            "month_hours": month_hours,
            "month_budget_hours": month_budget_hours,
            "month_percentage": month_percentage,
            "month_cost": month_cost,
            "month_budget_cost": month_budget_cost,
            "year_hours": year_hours,
            "year_budget_hours": year_budget_hours,
            "year_percentage": year_percentage,
            "expected_percentage": expected_percentage,
            "percentage_ratio": percentage_ratio,
            "year_cost": year_cost,
            "year_budget_cost": year_budget_cost,
            "month_color": month_color,
            "year_color": year_color,
            "horas_migradas": horas_migradas,
            "custo_migrado": custo_migrado,
            "proporcao_mes": proporcao_mes,
            "proporcao_atual": proporcao_atual
        })
    
    # Converter para DataFrame
    indicators_df = pd.DataFrame(project_indicators)
    
    # Layout de indicadores
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Meta Mensal")
        
        # Gráfico de meta mensal com cores e rótulos
        if not indicators_df.empty:
            fig_month = px.bar(
                indicators_df.sort_values("month_percentage", ascending=True),
                y="project_name",
                x="month_percentage",
                title=f"Percentual da Meta Mensal por Projeto - {calendar.month_name[month]}/{year}",
                labels={"project_name": "Projeto", "month_percentage": "% da Meta Mensal"},
                color_discrete_map={"green": "#4CAF50", "yellow": "#FFC107", "red": "#F44336"},
                color="month_color",
                orientation="h",
                text_auto='.1f'  # Adicionar rótulos automáticos com uma casa decimal
            )
            
            # Adicionar linhas de referência para os limites com anotações
            fig_month.add_shape(
                type="line",
                x0=60,
                x1=60,
                y0=-0.5,
                y1=len(indicators_df) - 0.5,
                line=dict(color="black", width=1, dash="dot"),
            )
            
            fig_month.add_annotation(
                x=60,
                y=len(indicators_df)/2,
                text="60%",
                showarrow=False,
                textangle=90,
                xshift=-15,
                font=dict(size=10, color="black")
            )
            
            fig_month.add_shape(
                type="line",
                x0=80,
                x1=80,
                y0=-0.5,
                y1=len(indicators_df) - 0.5,
                line=dict(color="black", width=1, dash="dot"),
            )
            
            fig_month.add_annotation(
                x=80,
                y=len(indicators_df)/2,
                text="80%",
                showarrow=False,
                textangle=90,
                xshift=-15,
                font=dict(size=10, color="black")
            )
            
            # Ajustar configurações do layout
            fig_month.update_layout(
                height=500,
                uniformtext_minsize=8,  # Tamanho mínimo do texto dos rótulos
                uniformtext_mode='hide'  # Esconder rótulos muito pequenos
            )
            
            # Ajustar posição dos rótulos
            fig_month.update_traces(textposition='auto')
            
            st.plotly_chart(fig_month, use_container_width=True)
        else:
            st.info("Não há dados suficientes para gerar o gráfico de meta mensal.")
    
    with col2:
        st.markdown("### Meta Anual")
        
        # Gráfico de meta anual com cores e rótulos - mantendo a visualização original
        if not indicators_df.empty:
            fig_year = px.bar(
                indicators_df.sort_values("year_percentage", ascending=True),
                y="project_name",
                x="year_percentage",
                title=f"Percentual da Meta Anual por Projeto - {year}",
                labels={"project_name": "Projeto", "year_percentage": "% da Meta Anual"},
                color_discrete_map={"green": "#4CAF50", "yellow": "#FFC107", "red": "#F44336"},
                color="year_color",
                orientation="h",
                text_auto='.1f'  # Adicionar rótulos automáticos com uma casa decimal
            )
            
            # Adicionar linhas de referência para os limites com anotações
            fig_year.add_shape(
                type="line",
                x0=40,
                x1=40,
                y0=-0.5,
                y1=len(indicators_df) - 0.5,
                line=dict(color="black", width=1, dash="dot"),
            )
            
            fig_year.add_annotation(
                x=40,
                y=len(indicators_df)/2,
                text="40%",
                showarrow=False,
                textangle=90,
                xshift=-15,
                font=dict(size=10, color="black")
            )
            
            fig_year.add_shape(
                type="line",
                x0=60,
                x1=60,
                y0=-0.5,
                y1=len(indicators_df) - 0.5,
                line=dict(color="black", width=1, dash="dot"),
            )
            
            fig_year.add_annotation(
                x=60,
                y=len(indicators_df)/2,
                text="60%",
                showarrow=False,
                textangle=90,
                xshift=-15,
                font=dict(size=10, color="black")
            )
            
            # Ajustar configurações do layout
            fig_year.update_layout(
                height=500,
                uniformtext_minsize=8,  # Tamanho mínimo do texto dos rótulos
                uniformtext_mode='hide'  # Esconder rótulos muito pequenos
            )
            
            # Ajustar posição dos rótulos
            fig_year.update_traces(textposition='auto')
            
            st.plotly_chart(fig_year, use_container_width=True)
        else:
            st.info("Não há dados suficientes para gerar o gráfico de meta anual.")
    
    # Tabela de indicadores com formatação melhorada
    st.markdown("### Indicadores Detalhados")
    
    # Definir estilo CSS para tornar a tabela mais atraente
    table_style = """
    <style>
        .styled-table {
            width: 100%;
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 14px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
        }
        .styled-table thead tr {
            background-color: #2c3e50;
            color: white;
            text-align: left;
            font-weight: bold;
        }
        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
            text-align: left;
        }
        .styled-table th:first-child,
        .styled-table td:first-child {
            text-align: left;
        }
        .styled-table tbody tr {
            border-bottom: 1px solid #dddddd;
        }
        .styled-table tbody tr:nth-of-type(even) {
            background-color: #f3f3f3;
        }
        .styled-table tbody tr:last-of-type {
            border-bottom: 2px solid #2c3e50;
        }
        .indicator-dot {
            border-radius: 50%;
            width: 15px;
            height: 15px;
            display: inline-block;
            margin-right: 6px;
            vertical-align: middle;
        }
        .number-cell {
            font-family: 'Courier New', monospace;
            font-weight: bold;
        }
    </style>
    """
    
    # Função auxiliar para converter horas decimais para formato HH:MM
    def decimal_to_hhmm(hours):
        if pd.isna(hours) or hours < 0:
            return "00:00"
        
        hours_int = int(hours)
        minutes_int = int((hours - hours_int) * 60)
        
        return f"{hours_int:02d}:{minutes_int:02d}"
    
    # Criar colunas estilizadas para os indicadores
    if not indicators_df.empty:
        display_df = indicators_df.copy()
        
        # Formatação dos valores
        display_df['Projeto'] = display_df['project_name']
        display_df['Cliente'] = display_df['client_name']
        display_df['Tipo'] = display_df['project_type']
        display_df['Status'] = display_df['status'].apply(lambda x: "Ativo" if str(x).lower() == "active" else "Inativo")
        display_df['Data Início'] = display_df['start_date']
        display_df['Data Fim'] = display_df['end_date']
        
        # Horas planejadas e realizadas (mensal) - formato HH:MM
        display_df['Horas Planeadas (Mês)'] = display_df['month_budget_hours'].apply(lambda x: f'<span class="number-cell">{decimal_to_hhmm(x)}</span>')
        display_df['Horas Realizadas (Mês)'] = display_df['month_hours'].apply(lambda x: f'<span class="number-cell">{decimal_to_hhmm(x)}</span>')
        
        # Custos planejados e realizados (mensal)
        display_df['Custo Planeado (Mês)'] = display_df['month_budget_cost'].apply(lambda x: f'<span class="number-cell">€{x:.2f}</span>')
        display_df['Custo Realizado (Mês)'] = display_df['month_cost'].apply(lambda x: f'<span class="number-cell">€{x:.2f}</span>')
        
        # Percentuais de conclusão
        display_df['Conclusão Esperada'] = display_df['expected_percentage'].apply(lambda x: f'<span class="number-cell">{x:.1f}%</span>')
        display_df['Conclusão Atual'] = display_df['year_percentage'].apply(lambda x: f'<span class="number-cell">{x:.1f}%</span>')
        
        # Custo total planejado e realizado
        display_df['Custo Total Planeado'] = display_df['year_budget_cost'].apply(lambda x: f'<span class="number-cell">€{x:.2f}</span>')
        display_df['Custo Total Realizado'] = display_df['year_cost'].apply(lambda x: f'<span class="number-cell">€{x:.2f}</span>')
        
        # Indicadores com cores - mantendo os indicadores originais
        display_df['Meta Mensal'] = display_df.apply(
            lambda x: f'<div><span class="indicator-dot" style="background-color: {"#4CAF50" if x.month_color == "green" else "#FFC107" if x.month_color == "yellow" else "#F44336"}"></span><span class="number-cell">{x.month_percentage:.1f}%</span></div>',
            axis=1
        )
        
        display_df['Meta Anual'] = display_df.apply(
            lambda x: f'<div><span class="indicator-dot" style="background-color: {"#4CAF50" if x.year_color == "green" else "#FFC107" if x.year_color == "yellow" else "#F44336"}"></span><span class="number-cell">{x.year_percentage:.1f}%</span></div>',
            axis=1
        )
        
        # Adicionar relação realizado/esperado como informação complementar
        display_df['Rel. Real/Esperado'] = display_df.apply(
            lambda x: f'<span class="number-cell">{x.percentage_ratio:.2f}</span>',
            axis=1
        )
        
        # Colunas para a tabela final
        table_columns = [
            'Projeto', 'Cliente', 'Tipo', 'Status', 'Data Início', 'Data Fim',
            'Horas Planeadas (Mês)', 'Horas Realizadas (Mês)',
            'Custo Planeado (Mês)', 'Custo Realizado (Mês)',
            'Conclusão Esperada', 'Conclusão Atual',
            'Custo Total Planeado', 'Custo Total Realizado',
            'Meta Mensal', 'Meta Anual', 'Rel. Real/Esperado'
        ]
        
        # Mostrar a tabela com as novas colunas e estilos
        st.markdown(table_style, unsafe_allow_html=True)
        st.write(
            '<table class="styled-table">' +
            display_df[table_columns].to_html(escape=False, index=False)
            .replace('<table', '<table class="styled-table"'),
            unsafe_allow_html=True
        )
        
        # Opção para download da tabela
        st.markdown("### Download dos Dados")
        
        # Preparar dados para download (sem formatação HTML)
        export_df = indicators_df.copy()
        
        # Formatar colunas para o arquivo de download
        export_df['Status'] = export_df['status'].apply(lambda x: "Ativo" if str(x).lower() == "active" else "Inativo")
        
        # Função auxiliar para converter horas decimais para formato HH:MM
        def decimal_to_hhmm_export(hours):
            if pd.isna(hours) or hours < 0:
                return "00:00"
            
            hours_int = int(hours)
            minutes_int = int((hours - hours_int) * 60)
            
            return f"{hours_int:02d}:{minutes_int:02d}"
        
        # Aplicar formatação para o export
        export_df['month_budget_hours_fmt'] = export_df['month_budget_hours'].apply(decimal_to_hhmm_export)
        export_df['month_hours_fmt'] = export_df['month_hours'].apply(decimal_to_hhmm_export)
        
        # Renomear colunas para o export
        export_df = export_df.rename(columns={
            'project_name': 'Projeto',
            'client_name': 'Cliente',
            'project_type': 'Tipo',
            'start_date': 'Data Início',
            'end_date': 'Data Fim',
            'month_budget_hours_fmt': 'Horas Planeadas (Mês)',
            'month_hours_fmt': 'Horas Realizadas (Mês)',
            'month_budget_cost': 'Custo Planeado (Mês)',
            'month_cost': 'Custo Realizado (Mês)',
            'year_budget_cost': 'Custo Total Planeado',
            'year_cost': 'Custo Total Realizado',
            'month_percentage': 'Meta Mensal (%)',
            'year_percentage': 'Meta Anual (%)',
            'expected_percentage': 'Conclusão Esperada (%)',
            'percentage_ratio': 'Rel. Real/Esperado'
        })
        
        # Selecionar colunas para exportação
        export_columns = [
            'Projeto', 'Cliente', 'Tipo', 'Status', 'Data Início', 'Data Fim',
            'Horas Planeadas (Mês)', 'Horas Realizadas (Mês)',
            'Custo Planeado (Mês)', 'Custo Realizado (Mês)',
            'Conclusão Esperada (%)', 'Meta Mensal (%)',
            'Custo Total Planeado', 'Custo Total Realizado',
            'Meta Anual (%)', 'Rel. Real/Esperado'
        ]
        
        # Criar buffer para o Excel
        output = io.BytesIO()
        
        # Opções de download
        col1, col2 = st.columns(2)
        
        with col1:
            # Botão para baixar Excel
            if st.button("📊 Baixar como Excel", type="primary"):
                # Criar arquivo Excel
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    export_df[export_columns].to_excel(writer, sheet_name='Indicadores de Projetos', index=False)
                    
                    # Acesso ao objeto workbook e worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Indicadores de Projetos']
                    
                    # Formatar cabeçalhos
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'fg_color': '#2c3e50',
                        'font_color': 'white',
                        'border': 1
                    })
                    
                    # Formato para valores numéricos com 2 casas decimais
                    number_format = workbook.add_format({'num_format': '0.00'})
                    
                    # Formato para valores monetários
                    money_format = workbook.add_format({'num_format': '€#,##0.00'})
                    
                    # Formato para percentuais
                    percent_format = workbook.add_format({'num_format': '0.00%'})
                    
                    # Aplicar formatação aos cabeçalhos
                    for col_num, value in enumerate(export_df[export_columns].columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        
                        # Aplicar formatos específicos às colunas
                        if 'Custo' in value:
                            worksheet.set_column(col_num, col_num, 15, money_format)
                        elif ('Meta' in value or 'Conclusão' in value) and '%' in value:
                            # Converter percentuais para valores decimais para formatação
                            for row_num, x in enumerate(export_df[value], start=1):
                                try:
                                    # Remover o símbolo % se presente e converter para decimal
                                    if isinstance(x, str) and '%' in x:
                                        x = float(x.replace('%', '')) / 100
                                    elif isinstance(x, (int, float)):
                                        x = x / 100
                                    worksheet.write(row_num, col_num, x, percent_format)
                                except:
                                    worksheet.write(row_num, col_num, x)
                            continue  # Pular a configuração padrão de coluna abaixo
                        else:
                            worksheet.set_column(col_num, col_num, 15)
                    
                    # Definir larguras de colunas específicas
                    worksheet.set_column('A:A', 25)  # Projeto
                    worksheet.set_column('B:B', 25)  # Cliente
                    worksheet.set_column('C:C', 15)  # Tipo
                
                # Oferecer o download
                output.seek(0)
                
                # Criar nome do arquivo com data atual
                current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name = f"indicadores_projetos_{current_date}.xlsx"
                
                st.download_button(
                    label="📥 Baixar Excel",
                    data=output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col2:
            # Botão para baixar CSV
            if st.button("📄 Baixar como CSV"):
                # Converter DataFrame para CSV
                csv = export_df[export_columns].to_csv(index=False)
                
                # Criar nome do arquivo com data atual
                current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name = f"indicadores_projetos_{current_date}.csv"
                
                st.download_button(
                    label="📥 Baixar CSV",
                    data=csv,
                    file_name=file_name,
                    mime="text/csv"
                )
        
        # Informações e legenda de cores
        st.markdown("### Informações e Legenda")
        
        # Criar um card informativo sobre o mês de referência
        st.markdown(f"""
        <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 15px;">
            <h4 style="margin-top: 0;">Resumo do Mês</h4>
            <p><strong>Mês de Referência:</strong> {month_name}/{year}</p>
            <p><strong>Equipa:</strong> {selected_team}</p>
            <p><strong>Projetos Analisados:</strong> {len(indicators_df)}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Legenda para meta mensal
        st.markdown("#### Meta Mensal")
        st.markdown("""
        <div style="display: flex; margin-top: 15px;">
            <div style="flex: 1; padding: 10px; margin-right: 10px; background-color: rgba(76, 175, 80, 0.1); border-left: 4px solid #4CAF50; border-radius: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="background-color: #4CAF50; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                    <strong>< 60% da meta</strong> - Consumo saudável de horas
                </div>
            </div>
            <div style="flex: 1; padding: 10px; margin-right: 10px; background-color: rgba(255, 193, 7, 0.1); border-left: 4px solid #FFC107; border-radius: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="background-color: #FFC107; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                    <strong>Entre 60% e 80% da meta</strong> - Monitorar consumo de horas
                </div>
            </div>
            <div style="flex: 1; padding: 10px; background-color: rgba(244, 67, 54, 0.1); border-left: 4px solid #F44336; border-radius: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="background-color: #F44336; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                    <strong>> 80% da meta</strong> - Consumo acelerado, atenção necessária
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Legenda para meta anual
        st.markdown("#### Meta Anual")
        st.markdown("""
        <div style="display: flex; margin-top: 15px;">
            <div style="flex: 1; padding: 10px; margin-right: 10px; background-color: rgba(76, 175, 80, 0.1); border-left: 4px solid #4CAF50; border-radius: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="background-color: #4CAF50; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                    <strong>< 40% da meta</strong> - Projeto com bom desempenho anual
                </div>
            </div>
            <div style="flex: 1; padding: 10px; margin-right: 10px; background-color: rgba(255, 193, 7, 0.1); border-left: 4px solid #FFC107; border-radius: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="background-color: #FFC107; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                    <strong>Entre 40% e 60% da meta</strong> - Projeto em atenção
                </div>
            </div>
            <div style="flex: 1; padding: 10px; background-color: rgba(244, 67, 54, 0.1); border-left: 4px solid #F44336; border-radius: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="background-color: #F44336; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
                    <strong>> 60% da meta</strong> - Projeto requer intervenção
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Explicação sobre a relação entre realizado e esperado
        st.markdown("#### Significado da Relação Realizado/Esperado")
        st.markdown("""
        <div style="background-color: #f0f7ff; padding: 15px; border-radius: 5px; margin-top: 15px; border-left: 4px solid #1976D2;">
            <p>O valor <strong>Rel. Real/Esperado</strong> indica a proporção entre o percentual concluído do projeto e o percentual que 
            deveria estar concluído de acordo com o tempo decorrido:</p>
            <ul>
                <li><strong>Valor = 1.0:</strong> O projeto está consumindo horas exatamente conforme esperado pelo tempo decorrido.</li>
                <li><strong>Valor < 1.0:</strong> O projeto está consumindo menos horas que o esperado para o tempo decorrido (positivo).</li>
                <li><strong>Valor > 1.0:</strong> O projeto está consumindo mais horas que o esperado para o tempo decorrido (negativo).</li>
            </ul>
            <p>Este é um indicador complementar à meta anual, que ajuda a entender se o projeto está no caminho certo considerando seu cronograma.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Explicação detalhada da metodologia
        with st.expander("Metodologia de Cálculos"):
            st.markdown("""
            #### Metodologia de Cálculo dos Indicadores
            
            **Meta Mensal** - Calculada com base na proporção do projeto no mês atual:
            
            1. **Proporção do Mês**: Para cada projeto, calculamos qual a proporção de dias úteis do projeto que deveriam ser executados no mês de referência.
            2. **Meta Mensal de Horas**: Total de horas do projeto × Proporção do mês.
            3. **Meta Mensal de Custo**: Total de custo do projeto × Proporção do mês.
            4. **Percentual de Conclusão Mensal**: (Horas realizadas no mês / Meta mensal de horas) × 100.
            
            **Meta Anual** - Mantida como no sistema original:
            
            1. **Percentual de Conclusão Anual**: (Horas realizadas acumuladas no ano / Total de horas planejadas) × 100.
            2. A classificação por cores permanece: <40% (verde), 40-60% (amarelo), >60% (vermelho).
            
            **Indicadores Complementares Adicionados**:
            
            1. **Conclusão Esperada**: Percentual do projeto que deveria estar concluído com base no tempo decorrido.
            2. **Conclusão Atual**: Percentual das horas planejadas que já foram efetivamente realizadas.
            3. **Relação Realizado/Esperado**: Conclusão Atual / Conclusão Esperada (indica se o projeto está consumindo recursos mais rápido ou mais lento que o planejado).
            
            **Observações importantes**:
            
            - As horas extras são contabilizadas como dobro para os cálculos (2× as horas normais).
            - Dados migrados de sistemas anteriores estão incluídos nos totais acumulados.
            - Os indicadores consideram apenas dias úteis (seg-sex, excluindo feriados).
            """)
    else:
        st.warning("Não há dados disponíveis para exibir na tabela.")

def show_revenue_indicators(db_manager, annual_target_manager, billing_manager, month, year, selected_team):
    """
    Mostra indicadores de faturação baseados nos registros reais de faturas
    """
    st.subheader("Indicadores de Faturação")
    
    # Carregar dados necessários
    annual_targets = annual_target_manager.read()
    invoices_df = billing_manager.get_invoice()  # Buscar todas as faturas
    clients_df = db_manager.query_to_df("SELECT * FROM clients")
    projects_df = db_manager.query_to_df("SELECT * FROM projects")
    groups_df = db_manager.query_to_df("SELECT * FROM groups")
    
    # Filtrar por equipe, se necessário
    if selected_team != "Todas":
        annual_targets = annual_targets[annual_targets["company_name"] == selected_team]
    
    if annual_targets.empty:
        st.warning("Não foram encontradas metas de faturação com os filtros selecionados.")
        return
    
    # Calcular as metas para o ano atual
    year_targets = annual_targets[annual_targets["target_year"] == year]
    
    if year_targets.empty:
        st.warning(f"Não foram encontradas metas definidas para o ano {year}.")
        return
    
    # Verificar se temos dados de faturação
    if invoices_df.empty:
        st.warning("Não há registros de faturação disponíveis. Por favor, registre faturas no sistema.")
        
        # Mostrar instruções sobre como registrar faturas
        with st.expander("Como registrar faturas"):
            st.markdown("""
            #### Como registrar faturas no sistema
            
            1. Acesse o menu **Gestão de Faturas** no sistema
            2. Na aba **Registar Faturas**, selecione o cliente e o projeto
            3. Preencha os detalhes da fatura: número, valor, datas, etc.
            4. Clique em **Registar Fatura** para salvar
            
            As faturas registradas serão automaticamente consideradas nos indicadores de faturação.
            """)
    
    # Definir períodos
    primeiro_dia_mes = datetime(year, month, 1)
    ultimo_dia_mes = datetime(year, month, calendar.monthrange(year, month)[1], 23, 59, 59)
    
    inicio_ano = datetime(year, 1, 1)
    fim_ano = datetime(year, 12, 31, 23, 59, 59)
    
    # Determinar trimestre atual
    current_quarter = (month - 1) // 3 + 1
    inicio_trimestre = datetime(year, ((current_quarter - 1) * 3) + 1, 1)
    fim_trimestre_mes = min(current_quarter * 3, 12)
    ultimo_dia_trimestre = calendar.monthrange(year, fim_trimestre_mes)[1]
    fim_trimestre = datetime(year, fim_trimestre_mes, ultimo_dia_trimestre, 23, 59, 59)
    
    # Calcular indicadores de faturação
    revenue_indicators = []
    
    for _, target in year_targets.iterrows():
        company_name = target["company_name"]
        annual_target = target["target_value"]
        
        # Calcular metas mensais e trimestrais
        monthly_target = annual_target / 12
        quarterly_target = annual_target / 4
        
        # Filtrar faturas por empresa usando projetos associados à empresa
        company_invoices = invoices_df.copy()
        
        if company_name != "Todas" and not invoices_df.empty:
            # Mapear projetos da empresa
            company_projects = []
            
            # Identificar o group_id correspondente à empresa
            group_id = None
            for _, group in groups_df.iterrows():
                if group['group_name'] == company_name:
                    group_id = group['id']
                    break
            
            if group_id is not None:
                # Filtrar projetos da empresa pelo group_id
                company_projects_df = projects_df[projects_df['group_id'] == group_id]
                company_projects = company_projects_df['project_id'].tolist()
                
                # Filtrar faturas pelos projetos da empresa
                if company_projects:
                    company_invoices = company_invoices[company_invoices['project_id'].isin(company_projects)]
                else:
                    company_invoices = pd.DataFrame()  # Sem projetos, sem faturas
        
        # Calcular faturação efetiva por período usando datas de pagamento das faturas
        if not company_invoices.empty:
            # Converter datas para datetime
            company_invoices['payment_date'] = pd.to_datetime(company_invoices['payment_date'])
            
            # Filtrar para o mês atual
            month_invoices = company_invoices[
                (company_invoices['payment_date'] >= primeiro_dia_mes) & 
                (company_invoices['payment_date'] <= ultimo_dia_mes)
            ]
            monthly_revenue = month_invoices['amount'].sum() if not month_invoices.empty else 0
            
            # Filtrar para o trimestre atual
            quarter_invoices = company_invoices[
                (company_invoices['payment_date'] >= inicio_trimestre) & 
                (company_invoices['payment_date'] <= fim_trimestre)
            ]
            quarterly_revenue = quarter_invoices['amount'].sum() if not quarter_invoices.empty else 0
            
            # Filtrar para o ano atual
            year_invoices = company_invoices[
                (company_invoices['payment_date'] >= inicio_ano) & 
                (company_invoices['payment_date'] <= fim_ano)
            ]
            annual_revenue = year_invoices['amount'].sum() if not year_invoices.empty else 0
        else:
            # Sem faturas registradas
            monthly_revenue = 0
            quarterly_revenue = 0
            annual_revenue = 0
        
        # Calcular percentuais de conclusão
        monthly_percentage = (monthly_revenue / monthly_target * 100) if monthly_target > 0 else 0
        quarterly_percentage = (quarterly_revenue / quarterly_target * 100) if quarterly_target > 0 else 0
        annual_percentage = (annual_revenue / annual_target * 100) if annual_target > 0 else 0
        
        # Determinar cores dos indicadores
        # Mensal
        if monthly_percentage >= 80:
            monthly_color = "green"
        elif monthly_percentage >= 60:
            monthly_color = "yellow"
        else:
            monthly_color = "red"
        
        # Trimestral
        if quarterly_percentage >= 80:
            quarterly_color = "green"
        elif quarterly_percentage >= 60:
            quarterly_color = "yellow"
        else:
            quarterly_color = "red"
            
        # Anual
        if annual_percentage >= 80:
            annual_color = "green"
        elif annual_percentage >= 60:
            annual_color = "yellow"
        else:
            annual_color = "red"
        
        # Adicionar ao array de indicadores
        revenue_indicators.append({
            "company_name": company_name,
            "monthly_target": monthly_target,
            "monthly_revenue": monthly_revenue,
            "monthly_percentage": monthly_percentage,
            "quarterly_target": quarterly_target,
            "quarterly_revenue": quarterly_revenue,
            "quarterly_percentage": quarterly_percentage,
            "annual_target": annual_target,
            "annual_revenue": annual_revenue,
            "annual_percentage": annual_percentage,
            "monthly_color": monthly_color,
            "quarterly_color": quarterly_color,
            "annual_color": annual_color
        })
    
    # Converter para DataFrame
    indicators_df = pd.DataFrame(revenue_indicators)
    
    # Layout de indicadores
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### Meta Mensal")
        
        # Gráfico de meta mensal com cores
        fig_month = px.bar(
            indicators_df,
            x="company_name",
            y="monthly_percentage",
            title=f"Percentual da Meta Mensal - {calendar.month_name[month]}/{year}",
            labels={"company_name": "Empresa", "monthly_percentage": "% da Meta Mensal"},
            color_discrete_map={"green": "#4CAF50", "yellow": "#FFC107", "red": "#F44336"},
            color="monthly_color",
            text_auto='.1f'
        )
        
        # Adicionar linhas de referência para os limites
        fig_month.add_shape(
            type="line",
            y0=60,
            y1=60,
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            line=dict(color="black", width=1, dash="dot"),
        )
        
        fig_month.add_shape(
            type="line",
            y0=80,
            y1=80,
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            line=dict(color="black", width=1, dash="dot"),
        )
        
        fig_month.update_layout(height=500)
        st.plotly_chart(fig_month, use_container_width=True)
    
    with col2:
        st.markdown("### Meta Trimestral")
        
        # Gráfico de meta trimestral com cores
        fig_quarter = px.bar(
            indicators_df,
            x="company_name",
            y="quarterly_percentage",
            title=f"Percentual da Meta Trimestral - {current_quarter}º Trimestre/{year}",
            labels={"company_name": "Empresa", "quarterly_percentage": "% da Meta Trimestral"},
            color_discrete_map={"green": "#4CAF50", "yellow": "#FFC107", "red": "#F44336"},
            color="quarterly_color",
            text_auto='.1f'
        )
        
        # Adicionar linhas de referência para os limites
        fig_quarter.add_shape(
            type="line",
            y0=60,
            y1=60,
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            line=dict(color="black", width=1, dash="dot"),
        )
        
        fig_quarter.add_shape(
            type="line",
            y0=80,
            y1=80,
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            line=dict(color="black", width=1, dash="dot"),
        )
        
        fig_quarter.update_layout(height=500)
        st.plotly_chart(fig_quarter, use_container_width=True)
        
    with col3:
        st.markdown("### Meta Anual")
        
        # Gráfico de meta anual com cores
        fig_year = px.bar(
            indicators_df,
            x="company_name",
            y="annual_percentage",
            title=f"Percentual da Meta Anual - {year}",
            labels={"company_name": "Empresa", "annual_percentage": "% da Meta Anual"},
            color_discrete_map={"green": "#4CAF50", "yellow": "#FFC107", "red": "#F44336"},
            color="annual_color",
            text_auto='.1f'
        )
        
        # Adicionar linhas de referência para os limites
        fig_year.add_shape(
            type="line",
            y0=60,
            y1=60,
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            line=dict(color="black", width=1, dash="dot"),
        )
        
        fig_year.add_shape(
            type="line",
            y0=80,
            y1=80,
            x0=-0.5,
            x1=len(indicators_df) - 0.5,
            line=dict(color="black", width=1, dash="dot"),
        )
        
        fig_year.update_layout(height=500)
        st.plotly_chart(fig_year, use_container_width=True)

    st.markdown("---")
    
    # Tabela de indicadores
    st.markdown("### Indicadores Detalhados")
    
    # Definir estilo CSS para a tabela (mesma definição das outras tabs)
    table_style = """
    <style>
        .styled-table {
            width: 100%;
            border-collapse: collapse;
            margin: 5px 0;
            font-size: 14px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
        }
        .styled-table thead tr {
            background-color: #2c3e50;
            color: white;
            text-align: left;
            font-weight: bold;
        }
        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
            text-align: left;
        }
        .styled-table th:first-child,
        .styled-table td:first-child {
            text-align: left;
        }
        .styled-table tbody tr {
            border-bottom: 1px solid #dddddd;
        }
        .styled-table tbody tr:nth-of-type(even) {
            background-color: #f3f3f3;
        }
        .styled-table tbody tr:last-of-type {
            border-bottom: 2px solid #2c3e50;
        }
        .indicator-dot {
            border-radius: 50%;
            width: 15px;
            height: 15px;
            display: inline-block;
            margin-right: 6px;
            vertical-align: middle;
        }
        .number-cell {
            font-family: 'Courier New', monospace;
            font-weight: bold;
        }
    </style>
    """
    
    st.markdown(table_style, unsafe_allow_html=True)
    
    # Função auxiliar para criar o indicador colorido
    def create_indicator(value, target, color):
        percentage = f"{value/target*100:.1f}%" if target > 0 else "0.0%"
        return f"""
        <div style="display: flex; align-items: center;">
            <div style="background-color: {color_map[color]}; border-radius: 50%; width: 15px; height: 15px; margin-right: 10px;"></div>
            <span class="number-cell">{percentage}</span>
            <span style="margin-left: 10px; color: #666;">(€{value:,.2f} / €{target:,.2f})</span>
        </div>
        """.strip()

    
    # Mapeamento de cores
    color_map = {
        "green": "#4CAF50", 
        "yellow": "#FFC107", 
        "red": "#F44336"
    }
    
    # Criar colunas estilizadas para os indicadores
    display_df = pd.DataFrame()
    display_df['Empresa'] = indicators_df['company_name']
    
    # Criar colunas de indicadores HTML
    display_df['Meta Mensal'] = indicators_df.apply(
    lambda x: create_indicator(x['monthly_revenue'], x['monthly_target'], x['monthly_color']).replace("\n", ""),
    axis=1
    )

    
    display_df['Meta Trimestral'] = indicators_df.apply(
        lambda x: create_indicator(x['quarterly_revenue'], x['quarterly_target'], x['quarterly_color']).replace("\n", ""),
        axis=1
    )
    
    display_df['Meta Anual'] = indicators_df.apply(
        lambda x: create_indicator(x['annual_revenue'], x['annual_target'], x['annual_color']).replace("\n", ""),
        axis=1
    )
    
    # Mostrar a tabela formatada
    st.write(
        '<table class="styled-table">' +
        display_df.to_html(escape=False, index=False)
        .replace('<table', '<table class="styled-table"'),
        unsafe_allow_html=True
    )

    # Legenda de cores
    st.markdown("### Legenda")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div style="display: flex; align-items: center;"><div style="background-color: #4CAF50; border-radius: 50%; width: 15px; height: 15px; margin-right: 5px;"></div> ≥ 80% da meta</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div style="display: flex; align-items: center;"><div style="background-color: #FFC107; border-radius: 50%; width: 15px; height: 15px; margin-right: 5px;"></div> Entre 60% e 80% da meta</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div style="display: flex; align-items: center;"><div style="background-color: #F44336; border-radius: 50%; width: 15px; height: 15px; margin-right: 5px;"></div> < 60% da meta</div>', unsafe_allow_html=True)
    
    # Mostrar detalhes das faturas registradas
    st.markdown("---")
    st.markdown("### Faturas Registradas")
    
    if not invoices_df.empty:
        # Faturas do ano atual
        year_invoices = invoices_df[
            pd.to_datetime(invoices_df['payment_date']).dt.year == year
        ]
        
        if not year_invoices.empty:
            # Seleção de visualização por período
            view_period = st.radio(
                "Visualizar faturas por:",
                ["Ano Completo", "Trimestre Atual", "Mês Atual"],
                horizontal=True
            )
            
            # Filtrar conforme período selecionado
            if view_period == "Mês Atual":
                filtered_invoices = year_invoices[
                    (pd.to_datetime(year_invoices['payment_date']).dt.month == month)
                ]
                period_desc = f"{calendar.month_name[month]}/{year}"
            elif view_period == "Trimestre Atual":
                filtered_invoices = year_invoices[
                    (pd.to_datetime(year_invoices['payment_date']).dt.month.between(
                        ((current_quarter - 1) * 3) + 1, 
                        current_quarter * 3
                    ))
                ]
                period_desc = f"{current_quarter}º Trimestre/{year}"
            else:  # Ano Completo
                filtered_invoices = year_invoices
                period_desc = str(year)
            
            if not filtered_invoices.empty:
                # Preparar para exibição
                display_invoices = filtered_invoices.copy()
                
                # Adicionar datas formatadas
                display_invoices['issue_date_fmt'] = pd.to_datetime(display_invoices['issue_date']).dt.strftime('%d/%m/%Y')
                display_invoices['payment_date_fmt'] = pd.to_datetime(display_invoices['payment_date']).dt.strftime('%d/%m/%Y')
                
                # Preparar colunas para exibição
                display_cols = [
                    'invoice_number', 'client_name', 'project_name', 
                    'amount', 'issue_date_fmt', 'payment_date_fmt'
                ]
                
                rename_cols = {
                    'invoice_number': 'Número da Fatura',
                    'client_name': 'Cliente',
                    'project_name': 'Projeto',
                    'amount': 'Valor',
                    'issue_date_fmt': 'Data de Emissão',
                    'payment_date_fmt': 'Data de Pagamento'
                }
                
                # Formatar valores monetários
                display_invoices['amount'] = display_invoices['amount'].apply(lambda x: f"€{float(x):,.2f}")
                
                # Mostrar resumo
                total_faturas = len(filtered_invoices)
                valor_total = filtered_invoices['amount'].sum()
                
                # Mostrar estatísticas
                st.info(f"**{total_faturas}** faturas registradas em **{period_desc}** - Total: **€{valor_total:,.2f}**")
                
                # Exibir dataframe
                st.dataframe(
                    display_invoices[display_cols].rename(columns=rename_cols),
                    use_container_width=True
                )
                
                # Gráfico de faturação ao longo do tempo
                st.markdown("#### Faturação ao longo do tempo")
                
                # Agrupar por mês para visualização
                if view_period == "Ano Completo":
                    # Agrupar por mês
                    invoices_by_month = filtered_invoices.copy()
                    invoices_by_month['month'] = pd.to_datetime(invoices_by_month['payment_date']).dt.month
                    monthly_totals = invoices_by_month.groupby('month')['amount'].sum().reset_index()
                    
                    # Adicionar nomes dos meses
                    monthly_totals['month_name'] = monthly_totals['month'].apply(lambda m: calendar.month_name[m])
                    
                    # Gráfico de barras por mês
                    fig_monthly = px.bar(
                        monthly_totals,
                        x='month_name',
                        y='amount',
                        title=f'Faturação Mensal em {year}',
                        labels={'month_name': 'Mês', 'amount': 'Valor Faturado (€)'},
                        text_auto='.2s'
                    )
                    
                    fig_monthly.update_layout(height=400)
                    st.plotly_chart(fig_monthly, use_container_width=True)
                    
                elif view_period == "Trimestre Atual":
                    # Se for trimestre, agrupar por semana
                    invoices_by_week = filtered_invoices.copy()
                    invoices_by_week['week'] = pd.to_datetime(invoices_by_week['payment_date']).dt.isocalendar().week
                    weekly_totals = invoices_by_week.groupby('week')['amount'].sum().reset_index()
                    
                    # Gráfico de barras por semana
                    fig_weekly = px.bar(
                        weekly_totals,
                        x='week',
                        y='amount',
                        title=f'Faturação Semanal no {current_quarter}º Trimestre de {year}',
                        labels={'week': 'Semana', 'amount': 'Valor Faturado (€)'},
                        text_auto='.2s'
                    )
                    
                    fig_weekly.update_layout(height=400)
                    st.plotly_chart(fig_weekly, use_container_width=True)
                
                else:  # Mês atual
                    # Se for mês, agrupar por dia
                    invoices_by_day = filtered_invoices.copy()
                    invoices_by_day['day'] = pd.to_datetime(invoices_by_day['payment_date']).dt.day
                    daily_totals = invoices_by_day.groupby('day')['amount'].sum().reset_index()
                    
                    # Gráfico de barras por dia
                    fig_daily = px.bar(
                        daily_totals,
                        x='day',
                        y='amount',
                        title=f'Faturação Diária em {calendar.month_name[month]} de {year}',
                        labels={'day': 'Dia', 'amount': 'Valor Faturado (€)'},
                        text_auto='.2s'
                    )
                    
                    fig_daily.update_layout(height=400)
                    st.plotly_chart(fig_daily, use_container_width=True)
            else:
                st.info(f"Não foram encontradas faturas para o período selecionado ({period_desc}).")
        else:
            st.info(f"Não foram encontradas faturas para o ano {year}.")
    else:
        st.info("Não há faturas registradas no sistema.")
    
    