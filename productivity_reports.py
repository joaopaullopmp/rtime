# productivity_reports.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
import plotly.graph_objects as go
from database_manager import DatabaseManager
from report_utils import calcular_dias_uteis_projeto, get_feriados_portugal

def calcular_usuarios_por_equipe(users_df):
    """Calcula o número de usuários ativos por equipe"""
    usuarios_por_equipe = {}
    
    for _, user in users_df[users_df['active'] == True].iterrows():
        grupos = eval(user['groups']) if isinstance(user['groups'], str) else []
        for grupo in grupos:
            if grupo not in usuarios_por_equipe:
                usuarios_por_equipe[grupo] = 0
            usuarios_por_equipe[grupo] += 1
    
    return usuarios_por_equipe

def calcular_horas_uteis_por_equipe(inicio_mes, fim_mes, users_df):
    """Calcula o total de horas úteis por equipe"""
    dias_uteis = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    horas_uteis_mes = dias_uteis * 8  # 8 horas por dia útil
    
    usuarios_por_equipe = calcular_usuarios_por_equipe(users_df)
    return {
        equipe: horas_uteis_mes * total_usuarios
        for equipe, total_usuarios in usuarios_por_equipe.items()
    }

def calcular_ausencias_equipe(absences_df, users_df, inicio_mes, fim_mes, equipe=None):
    """Calcula o percentual de ausências por equipe"""
    ausencias_por_equipe = {}
    
    # Filtrar ausências do período
    ausencias_periodo = absences_df[
    (pd.to_datetime(absences_df['start_date'], format='mixed') <= fim_mes) &
    (pd.to_datetime(absences_df['end_date'], format='mixed') >= inicio_mes)
]
    
    # Para cada usuário, calcular dias úteis de ausência
    for _, user in users_df[users_df['active'] == True].iterrows():
        grupos = eval(user['groups']) if isinstance(user['groups'], str) else []
        if equipe and equipe not in grupos:
            continue
            
        user_absences = ausencias_periodo[ausencias_periodo['user_id'] == user['user_id']]
        
        for grupo in grupos:
            if grupo not in ausencias_por_equipe:
                ausencias_por_equipe[grupo] = {'dias_ausencia': 0, 'total_usuarios': 0}
            ausencias_por_equipe[grupo]['total_usuarios'] += 1
            
            for _, absence in user_absences.iterrows():
                start = max(inicio_mes, pd.to_datetime(absence['start_date']))
                end = min(fim_mes, pd.to_datetime(absence['end_date']))
                dias_uteis = calcular_dias_uteis_projeto(start.date(), end.date())
                ausencias_por_equipe[grupo]['dias_ausencia'] += dias_uteis
    
    # Calcular percentual de ausências
    dias_uteis_mes = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    for equipe, dados in ausencias_por_equipe.items():
        total_dias_uteis = dias_uteis_mes * dados['total_usuarios']
        ausencias_por_equipe[equipe]['percentual'] = (
            dados['dias_ausencia'] / total_dias_uteis * 100 
            if total_dias_uteis > 0 else 0
        )
    
    return ausencias_por_equipe

def calcular_ausencias_usuario(absences_df, user_id, inicio_mes, fim_mes):
    """Calcula ausências para um usuário específico"""
    # Filtrar ausências do período para o usuário
    ausencias_periodo = absences_df[
    (absences_df['user_id'] == user_id) &
    (pd.to_datetime(absences_df['start_date'], format='mixed') <= fim_mes) &
    (pd.to_datetime(absences_df['end_date'], format='mixed') >= inicio_mes)
]
    
    dias_ausencia = 0
    if not ausencias_periodo.empty:
        for _, absence in ausencias_periodo.iterrows():
            start = max(inicio_mes, pd.to_datetime(absence['start_date']))
            end = min(fim_mes, pd.to_datetime(absence['end_date']))
            dias_uteis = calcular_dias_uteis_projeto(start.date(), end.date())
            dias_ausencia += dias_uteis
    
    # Calcular percentual de ausências
    dias_uteis_mes = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    percentual_ausencia = (dias_ausencia / dias_uteis_mes * 100) if dias_uteis_mes > 0 else 0
    
    return {
        'dias_ausencia': dias_ausencia,
        'percentual': percentual_ausencia
    }

def calcular_metricas_produtividade_atualizado(dados, group_by_column, horas_uteis_mes, users_df, absences_df, inicio_mes, fim_mes):
    """Calcula métricas de produtividade considerando ausências e horas úteis por equipe"""
    # Calcular horas úteis por equipe
    horas_uteis_equipe = calcular_horas_uteis_por_equipe(inicio_mes, fim_mes, users_df)
    
    # Calcular ausências por equipe
    ausencias = calcular_ausencias_equipe(absences_df, users_df, inicio_mes, fim_mes)
    
    if dados.empty:
        return pd.DataFrame()
    
    # Calcular horas faturáveis e extras
    dados['horas_faturaveis'] = dados.apply(lambda row: row['hours'] if row.get('billable', False) else 0, axis=1)
    dados['horas_extra'] = dados.apply(lambda row: row['hours'] if row.get('overtime', False) else 0, axis=1)

    # No método calcular_metricas_produtividade_atualizado
    # Agrupar dados por equipe
    metricas = dados.groupby(group_by_column).agg({
        'hours': 'sum',  # Soma das horas (total_horas)
        'id': 'count',   # Contagem de registros (total_registros)
        'horas_faturaveis': 'sum',
        'horas_extra': 'sum'
    }).reset_index()

    # Renomear colunas
    metricas.columns = [group_by_column, 'total_horas', 'total_registros', 'horas_faturaveis', 'horas_extra']
    
    # Adicionar total de utilizadores por equipe
    usuarios_por_equipe = calcular_usuarios_por_equipe(users_df)
    metricas['total_utilizadores'] = metricas[group_by_column].map(usuarios_por_equipe)

    # Adicionar percentual de ausências e ajustar horas úteis disponíveis
    metricas['percentual_ausencias'] = metricas[group_by_column].map(
        lambda x: ausencias.get(x, {}).get('percentual', 0)
    )
    
    # Calcular horas úteis totais e efetivas
    metricas['horas_uteis_totais'] = metricas[group_by_column].map(horas_uteis_equipe)
    metricas['horas_uteis_disponiveis'] = metricas.apply(
        lambda row: max(row['horas_uteis_totais'] * (1 - row['percentual_ausencias']/100) - row['total_horas'], 0),
        axis=1
    )

    # Recalcular percentuais baseados nas horas úteis efetivas
    metricas['percentual_faturavel'] = metricas.apply(
        lambda row: (row['horas_faturaveis'] / row['horas_uteis_totais'] * 100)
        if row['horas_uteis_disponiveis'] > 0 else 0,
        axis=1
    ).round(2)
    
    metricas['percentual_extra'] = metricas.apply(
        lambda row: (row['horas_extra'] / row['total_horas'] * 100)
        if row['total_horas'] > 0 else 0,
        axis=1
    ).round(2)
    
    metricas['percentual_ocupacao'] = metricas.apply(
        lambda row: (row['total_horas'] / row['horas_uteis_totais'] * 100)
        if row['horas_uteis_disponiveis'] > 0 else 0,
        axis=1
    ).round(2)

    return metricas

def calcular_metricas_produtividade_usuario(dados, absences_df, inicio_mes, fim_mes):
    """Calcula métricas de produtividade por usuário"""
    # Calcular dias úteis e horas úteis do mês
    dias_uteis = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    horas_uteis_mes = dias_uteis * 8  # 8 horas por dia útil
    
    if dados.empty:
        return pd.DataFrame()
        
    # Calcular horas faturáveis e extras
    dados['horas_faturaveis'] = dados.apply(lambda row: row['hours'] if row.get('billable', False) else 0, axis=1)
    dados['horas_extra'] = dados.apply(lambda row: row['hours'] if row.get('overtime', False) else 0, axis=1)

    # Agrupar dados por usuário
    metricas = dados.groupby(['user_id', 'nome_completo', 'group_name']).agg({
        'hours': 'sum',  # Soma das horas (total_horas)
        'id': 'count',   # Contagem de registros (total_registros)
        'horas_faturaveis': 'sum',
        'horas_extra': 'sum'
    }).reset_index()

    # Renomear colunas
    metricas.columns = ['user_id', 'nome_completo', 'group_name', 'total_horas', 'total_registros', 
                    'horas_faturaveis', 'horas_extra']

    # Calcular ausências
    def calcular_horas_ausencia(user_id):
        ausencias = calcular_ausencias_usuario(absences_df, user_id, inicio_mes, fim_mes)
        return ausencias['dias_ausencia'] * 8  # Converter dias em horas

    # Adicionar métricas de horas
    metricas['horas_ausencia'] = metricas['user_id'].apply(calcular_horas_ausencia)
    metricas['horas_uteis_totais'] = horas_uteis_mes  # Total fixo para todos
    metricas['horas_uteis_disponiveis'] = metricas['horas_uteis_totais'] - metricas['total_horas'] - metricas['horas_ausencia']
    
    # Calcular percentuais
    # % Ausência é sobre o total de horas úteis
    metricas['percentual_ausencias'] = (metricas['horas_ausencia'] / metricas['horas_uteis_totais'] * 100).round(2)
    
    # % Ocupação é horas trabalhadas sobre horas disponíveis
    metricas['percentual_ocupacao'] = metricas.apply(
        lambda row: (row['total_horas'] / row['horas_uteis_totais'] * 100)
        if row['horas_uteis_totais'] > 0 else 0,
        axis=1
    ).round(2)
    
    # % Faturável é horas faturadas sobre horas totais disponíveis
    metricas['percentual_faturavel'] = metricas.apply(
        lambda row: (row['horas_faturaveis'] / row['horas_uteis_totais'] * 100)
        if row['horas_uteis_totais'] > 0 else 0,
        axis=1
    ).round(2)
    
    # % Hora Extra é sobre horas totais disponíveis
    metricas['percentual_extra'] = metricas.apply(
        lambda row: (row['horas_extra'] / row['horas_uteis_totais'] * 100)
        if row['horas_uteis_totais'] > 0 else 0,
        axis=1
    ).round(2)

    return metricas

def team_productivity_page():
    """Página de relatório de produtividade por equipe"""
    st.title("Relatório de Produtividade por Equipa")
    
    # Carregar dados do banco SQLite em vez de arquivos Excel
    db_manager = DatabaseManager()
    
    # Consultar tabelas necessárias
    try:
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        absences_df = db_manager.query_to_df("SELECT * FROM absences")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
    except Exception as e:
        st.error(f"Erro ao carregar dados do banco: {str(e)}")
        return
    
    # Criar lista de anos em ordem decrescente
    ano_atual = datetime.now().year
    anos = list(range(ano_atual, 2023, -1))
    
    # Filtros
    with st.expander("Filtros", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox(
                "Mês",
                range(1, 13),
                format_func=lambda x: calendar.month_name[x],
                index=datetime.now().month - 1,
                key="mes_team"
            )
        with col2:
            ano = st.selectbox(
                "Ano",
                anos,
                index=0,
                key="ano_team"
            )
    
    # Calcular datas do período - MODIFICADO para incluir todo o último dia
    inicio_mes = datetime(ano, mes, 1)
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    fim_mes = datetime(ano, mes, ultimo_dia, 23, 59, 59)  # Set to end of last day
    
    # Calcular dias úteis do mês considerando feriados
    dias_uteis = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    horas_uteis = dias_uteis * 8  # 8 horas por dia útil
    st.info(f"Dias úteis no mês: {dias_uteis} dias ({horas_uteis}h por pessoa)")
    
    # Processar dados
    dados = timesheet_df.copy()
    
    # Aplicar filtros de data
    # Alteração na linha 277 do productivity_reports.py
    dados['start_date'] = pd.to_datetime(dados['start_date'], format='mixed')
    dados = dados[
        (dados['start_date'] >= inicio_mes) &
        (dados['start_date'] <= fim_mes)
    ]
    
    if not dados.empty:
        # Criar um mapeamento de id para group_name
        group_map = dict(zip(groups_df['id'], groups_df['group_name']))

        # Adicionar a coluna 'group_name' aos dados usando o mapeamento
        dados['group_name'] = dados['group_id'].map(group_map)

        # Verificar se a coluna 'group_name' está presente
        if 'group_name' in dados.columns:
            # Calcular horas úteis por equipe
            horas_uteis_equipe = calcular_horas_uteis_por_equipe(inicio_mes, fim_mes, users_df)
            
            # Calcular métricas com o novo método
            metricas = calcular_metricas_produtividade_atualizado(
                dados, 
                'group_name', 
                horas_uteis, 
                users_df,
                absences_df,
                inicio_mes,
                fim_mes
            )
            
            # Exibir resultados com todas as métricas
            st.subheader(f"Métricas por Equipe - Mês de Referência: {mes}/{ano}")
            st.dataframe(
                metricas.rename(columns={
                    'group_name': 'Equipa',
                    'total_utilizadores': 'Total Colaboradores',
                    'total_horas': 'Horas Realizadas',
                    'total_registros': 'Total de Registros',
                    'horas_faturaveis': 'Horas Faturáveis',
                    'horas_extra': 'Horas Extras',
                    'percentual_faturavel': '% Faturável',
                    'percentual_extra': '% Hora Extra',
                    'percentual_ocupacao': '% Ocupação',
                    'percentual_ausencias': '% Ausências',
                    'horas_uteis_totais': 'Horas Úteis Totais',
                    'horas_uteis_disponiveis': 'Horas Úteis Disponíveis'
                })[[
                    'Equipa',
                    'Total Colaboradores',
                    'Horas Úteis Totais',
                    'Horas Realizadas',
                    'Horas Úteis Disponíveis',
                    '% Ocupação',
                    'Horas Faturáveis',
                    '% Faturável',
                    'Horas Extras',
                    '% Hora Extra',
                    '% Ausências'               
                ]].style.format({
                    'Total Colaboradores': '{:.0f}',
                    'Horas Realizadas': '{:.2f}',
                    'Horas Faturáveis': '{:.2f}',
                    'Horas Extras': '{:.2f}',
                    '% Faturável': '{:.1f}%',
                    '% Hora Extra': '{:.1f}%',
                    '% Ocupação': '{:.1f}%',
                    '% Ausências': '{:.1f}%',
                    'Horas Úteis Totais': '{:.1f}',
                    'Horas Úteis Disponíveis': '{:.1f}'
                }).hide(axis='index')
            )
            
            # Gráfico de produtividade atualizado
            st.subheader(f"Gráfico de Produtividade por Equipe - Mês de Referência: {mes}/{ano}")
            st.write("Metas: Ocupação 75% | Faturável 62.5%")
            
            dados_grafico = metricas[['group_name', 'percentual_ocupacao', 'percentual_faturavel', 'percentual_ausencias']].to_dict('records')
            
            fig = go.Figure()

            # Adicionar barras para ocupação e faturável
            fig.add_trace(go.Bar(
                x=[d['group_name'] for d in dados_grafico],
                y=[d['percentual_ocupacao'] for d in dados_grafico],
                name='Ocupação',
                marker_color='#8884d8',
                hovertemplate='%{y:.1f}%'
            ))

            fig.add_trace(go.Bar(
                x=[d['group_name'] for d in dados_grafico],
                y=[d['percentual_faturavel'] for d in dados_grafico],
                name='Faturável',
                marker_color='#82ca9d',
                hovertemplate='%{y:.1f}%'
            ))

            # Adicionar linha para ausências
            fig.add_trace(go.Scatter(
                x=[d['group_name'] for d in dados_grafico],
                y=[d['percentual_ausencias'] for d in dados_grafico],
                name='Ausências',
                mode='lines+markers',
                line=dict(color='#ff7f7f', width=2),
                hovertemplate='%{y:.1f}%'
            ))

            # Adicionar linhas de meta usando add_hline
            fig.add_hline(
                y=75,
                line_dash="dash",
                line_color="red",
                line_width=2,
                annotation_text="Meta Ocupação (75%)",
                annotation_position="right",
                annotation=dict(
                    font=dict(color="red", size=12),
                    xanchor="left"
                )
            )

            fig.add_hline(
                y=62.5,
                line_dash="dash",
                line_color="orange",
                line_width=2,
                annotation_text="Meta Faturável (62.5%)",
                annotation_position="right",
                annotation=dict(
                    font=dict(color="orange", size=12),
                    xanchor="left"
                )
            )

            fig.update_layout(
                barmode='group',
                title=(f'Gráfico de Produtividade por Equipe - Mês de Referência: {mes}/{ano}'),
                xaxis_title='Equipe',
                yaxis_title='Percentual (%)',
                yaxis=dict(
                    range=[0, 100],
                    ticksuffix='%',
                    gridcolor='lightgray',
                    gridwidth=1
                ),
                xaxis=dict(
                    gridcolor='lightgray',
                    gridwidth=1
                ),
                legend=dict(
                    orientation='h',
                    yanchor='bottom',
                    y=1.02,
                    xanchor='right',
                    x=1
                ),
                height=500,
                plot_bgcolor='white',
                margin=dict(r=150)  # Margem direita para acomodar os textos das linhas
            )

            # Adicionar grade horizontal mais visível
            fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')

            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Não foi possível calcular as métricas por equipe. A coluna 'group_name' está ausente nos dados.")
    else:
        st.warning("Nenhum registro encontrado para o período selecionado.")

def user_productivity_page():
    """Página de relatório de produtividade por Colaborador"""
    st.title("Relatório de Produtividade por Colaborador")
    
    # Carregar dados do banco SQLite em vez de arquivos Excel
    db_manager = DatabaseManager()
    
    # Consultar tabelas necessárias
    try:
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        absences_df = db_manager.query_to_df("SELECT * FROM absences")
    except Exception as e:
        st.error(f"Erro ao carregar dados do banco: {str(e)}")
        return
    
    # Criar lista de anos em ordem decrescente
    ano_atual = datetime.now().year
    anos = list(range(ano_atual, 2023, -1))
    
    # Filtros
    with st.expander("Filtros", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox(
                "Mês",
                range(1, 13),
                format_func=lambda x: calendar.month_name[x],
                index=datetime.now().month - 1,
                key="mes_user"
            )
        with col2:
            ano = st.selectbox(
                "Ano",
                anos,
                index=0,
                key="ano_user"
            )
    
    # Calcular datas do período - MODIFICADO para incluir todo o último dia
    inicio_mes = datetime(ano, mes, 1)
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    fim_mes = datetime(ano, mes, ultimo_dia, 23, 59, 59)  # Set to end of last day
    
    # Calcular dias úteis do mês considerando feriados
    dias_uteis = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    horas_uteis = dias_uteis * 8  # 8 horas por dia útil
    st.info(f"Dias úteis no mês: {dias_uteis} dias ({horas_uteis}h por pessoa)")
    
    # Processar dados
    dados = timesheet_df.copy()
    
    if dados.empty:
        st.warning("Não existem registros de horas no sistema.")
        return
    
    # Aplicar filtros de data
    dados['start_date'] = pd.to_datetime(dados['start_date'], format='mixed')
    dados = dados[
        (dados['start_date'] >= inicio_mes) & 
        (dados['start_date'] <= fim_mes)
    ]
    
    if not dados.empty:
        # Adicionar informações das equipes
        if 'group_id' in dados.columns and 'id' in groups_df.columns and 'group_name' in groups_df.columns:
            group_map = dict(zip(groups_df['id'], groups_df['group_name']))
            dados['group_name'] = dados['group_id'].map(group_map)
        else:
            st.error("As colunas necessárias para o mapeamento de 'group_name' estão ausentes.")
            return
        
        # Adicionar informações dos usuários
        if 'user_id' in dados.columns and 'user_id' in users_df.columns:
            dados = dados.merge(
                users_df,
                on='user_id',
                how='left'
            )
        else:
            st.error("As colunas necessárias para o mapeamento de usuários estão ausentes.")
            return

        # Criar coluna de nome completo
        dados['nome_completo'] = dados['First_Name'] + ' ' + dados['Last_Name']
        
        # Filtro de equipe
        if 'group_name' in dados.columns:
            equipes = dados['group_name'].unique().tolist()
            equipe_selecionada = st.multiselect(
                "Filtrar por Equipe", 
                equipes,
                key="equipe_user"
            )
        else:
            st.warning("Nenhuma equipe foi encontrada nos registros.")
            return

        # Filtrar por equipe se selecionada
        if equipe_selecionada:
            dados = dados[dados['group_name'].isin(equipe_selecionada)]
            if dados.empty:
                st.warning("Não foram encontrados registros para a(s) equipe(s) selecionada(s) no período.")
                return
        
        # Calcular métricas por usuário
        metricas = calcular_metricas_produtividade_usuario(
            dados,
            absences_df,
            inicio_mes,
            fim_mes
        )
        
        if metricas.empty:
            st.warning("Não foi possível calcular as métricas de produtividade para o período selecionado.")
            return
        
         # Exibir resultados
        st.subheader(f"Métricas por Colaborador - Equipa: {equipe_selecionada} - Mês de Referência: {mes}/{ano}")
        st.dataframe(
            metricas[[
                'nome_completo',
                'group_name',
                'horas_uteis_totais',
                'total_horas',
                'horas_uteis_disponiveis',
                'percentual_ocupacao',
                'horas_faturaveis',
                'percentual_faturavel',
                'horas_extra',
                'percentual_extra',
                'percentual_ausencias'
            ]].rename(columns={
                'nome_completo': 'Colaborador',
                'group_name': 'Equipe',
                'horas_uteis_totais': 'Horas Úteis Totais',
                'total_horas': 'Horas Realizadas',
                'horas_uteis_disponiveis': 'Horas Úteis Disponíveis',
                'percentual_ocupacao': '% Ocupação',
                'horas_faturaveis': 'Horas Faturáveis',
                'percentual_faturavel': '% Faturável',
                'horas_extra': 'Horas Extras',
                'percentual_extra': '% Hora Extra',
                'percentual_ausencias': '% Ausências'
            }).style.format({
                'Horas Realizadas': '{:.2f}',
                'Horas Faturáveis': '{:.2f}',
                '% Faturável': '{:.1f}%',
                'Horas Extras': '{:.2f}',
                '% Hora Extra': '{:.1f}%',
                '% Ocupação': '{:.1f}%',
                '% Ausências': '{:.1f}%',
                'Horas Úteis Totais': '{:.1f}',
                'Horas Úteis Disponíveis': '{:.1f}'
            }).hide(axis='index')
        )
        
        # Gráfico de produtividade
        st.subheader(f"Gráfico de Produtividade por Colaborador - Equipa: {equipe_selecionada} - Mês de Referência: {mes}/{ano}")
        st.write("Metas: Ocupação 75% | Faturável 62.5%")
        
        dados_grafico = metricas[['nome_completo', 'percentual_ocupacao', 'percentual_faturavel', 'percentual_ausencias']].to_dict('records')
        
        # Gráfico de produtividade por colaborador
        fig = go.Figure()

        # Adicionar barras para ocupação e faturável
        fig.add_trace(go.Bar(
            x=[d['nome_completo'] for d in dados_grafico],
            y=[d['percentual_ocupacao'] for d in dados_grafico],
            name='Ocupação',
            marker_color='#8884d8',
            hovertemplate='%{y:.1f}%'
        ))

        fig.add_trace(go.Bar(
            x=[d['nome_completo'] for d in dados_grafico],
            y=[d['percentual_faturavel'] for d in dados_grafico],
            name='Faturável',
            marker_color='#82ca9d',
            hovertemplate='%{y:.1f}%'
        ))

        # Adicionar linha para ausências
        fig.add_trace(go.Scatter(
            x=[d['nome_completo'] for d in dados_grafico],
            y=[d['percentual_ausencias'] for d in dados_grafico],
            name='Ausências',
            mode='lines+markers',
            line=dict(color='#ff7f7f', width=2),
            hovertemplate='%{y:.1f}%'
       ))

        # Adicionar linhas de meta usando add_hline
        fig.add_hline(
            y=75,
            line_dash="dash",
            line_color="red",
            line_width=2,
            annotation_text="Meta Ocupação (75%)",
            annotation_position="right",
            annotation=dict(
                font=dict(color="red", size=12),
                xanchor="left"
            )
        )

        fig.add_hline(
            y=62.5,
            line_dash="dash",
            line_color="orange",
            line_width=2,
            annotation_text="Meta Faturável (62.5%)",
            annotation_position="right",
            annotation=dict(
                font=dict(color="orange", size=12),
                xanchor="left"
            )
        )

        fig.update_layout(
            barmode='group',
            title=(f'Gráfico de Produtividade por Colaborador - Equipa: {equipe_selecionada} - Mês de Referência: {mes}/{ano}'),
            xaxis_title='Colaborador',
            yaxis_title='Percentual (%)',
            yaxis=dict(
                range=[0, 100],
                ticksuffix='%',
                gridcolor='lightgray',
                gridwidth=1
            ),
            xaxis=dict(
                tickangle=-45,
                gridcolor='lightgray',
                gridwidth=1
            ),
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=1.02,
                xanchor='right',
                x=1
            ),
            height=500,
            plot_bgcolor='white',
            margin=dict(r=150, b=100)  # Margem direita e inferior ajustadas
        )

        # Adicionar grade horizontal mais visível
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Nenhum registro encontrado para o período selecionado.")

    # Exports
    __all__ = [
    'team_productivity_page',
    'user_productivity_page'
    ]