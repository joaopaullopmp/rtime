# productivity_dashboard.py
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import calendar

from timesheet import TimesheetManager
from productivity_reports import (
    calcular_dias_uteis_projeto, 
    calcular_usuarios_por_equipe, 
    calcular_metricas_produtividade_atualizado,
    calcular_metricas_produtividade_usuario
)
from risk_reports import calcular_risco_projeto
from report_utils import get_feriados_portugal

def calcular_risco_projeto_dashboard(custo_realizado, custo_planejado):
    """
    Calcula o risco do projeto com base no custo realizado vs planejado
    """
    try:
        # Calcular variação percentual
        if custo_planejado == 0:
            return {
                "nivel": "Sem Dados",
                "cor": "gray",
                "recomendacao": "Impossível calcular risco",
                "variacao_percentual": 0
            }
        
        variacao_percentual = ((custo_realizado - custo_planejado) / custo_planejado) * 100

        # Definir níveis de risco baseados na variação
        if variacao_percentual <= -10:
            return {
                "nivel": "Eficiente",
                "cor": "green",
                "recomendacao": "Projeto abaixo do orçamento. Bom desempenho financeiro.",
                "variacao_percentual": variacao_percentual
            }
        elif -10 < variacao_percentual <= 10:
            return {
                "nivel": "No Planejado",
                "cor": "blue",
                "recomendacao": "Projeto dentro do orçamento. Desempenho financeiro adequado.",
                "variacao_percentual": variacao_percentual
            }
        elif 10 < variacao_percentual <= 20:
            return {
                "nivel": "Atenção",
                "cor": "orange",
                "recomendacao": "Projeto começando a ultrapassar o orçamento. Revisar escopo e custos.",
                "variacao_percentual": variacao_percentual
            }
        elif 20 < variacao_percentual <= 30:
            return {
                "nivel": "Risco Moderado",
                "cor": "red",
                "recomendacao": "Projeto significativamente acima do custo. Ações urgentes:\n- Revisar escopo\n- Identificar causas de sobrecusto",
                "variacao_percentual": variacao_percentual
            }
        else:
            return {
                "nivel": "Alto Risco",
                "cor": "darkred", 
                "recomendacao": "CRÍTICO: Projeto muito acima do custo. Medidas imediatas:\n- Suspender projeto\n- Renegociar contrato\n- Replanejar completamente",
                "variacao_percentual": variacao_percentual
            }
    except Exception as e:
        st.error(f"Erro ao calcular risco do projeto: {e}")
        return {
            "nivel": "Erro",
            "cor": "gray",
            "recomendacao": "Não foi possível calcular o risco",
            "variacao_percentual": 0
        }

def calcular_indicadores_consolidados(ano, mes):
    """
    Calcula indicadores consolidados de produtividade para o sistema
    """
    # Carregar dados necessários
    timesheet_df = pd.read_excel('timesheet.xlsx')
    users_df = pd.read_excel('utilizadores.xlsx')
    groups_df = pd.read_excel('groups.xlsx')
    projects_df = pd.read_excel('projects.xlsx')
    absences_df = pd.read_excel('absences.xlsx')
    rates_df = pd.read_excel('rates.xlsx')
    
    # Definir período
    inicio_mes = datetime(ano, mes, 1)
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    fim_mes = datetime(ano, mes, ultimo_dia)
    
    # Calcular dias úteis
    dias_uteis = calcular_dias_uteis_projeto(inicio_mes.date(), fim_mes.date())
    
    # Processar dados do timesheet
    dados = timesheet_df.copy()
    dados['start_date'] = pd.to_datetime(dados['start_date'])
    dados = dados[
        (dados['start_date'] >= inicio_mes) &
        (dados['start_date'] <= fim_mes)
    ]
    
    # Adicionar grupo ao timesheet
    dados = dados.merge(
        groups_df[['id', 'group_name']],
        left_on='group_id',
        right_on='id',
        how='left'
    )
    
    # Adicionar nome completo
    dados = dados.merge(
        users_df[['user_id', 'First_Name', 'Last_Name']],
        on='user_id',
        how='left'
    )
    dados['nome_completo'] = dados['First_Name'] + ' ' + dados['Last_Name']
    
    # Métricas de produtividade por equipe
    try:
        metricas_equipe = calcular_metricas_produtividade_atualizado(
            dados, 
            'group_name', 
            dias_uteis * 8,  # horas úteis 
            users_df,
            absences_df,
            inicio_mes,
            fim_mes
        )
    except Exception as e:
        st.error(f"Erro ao calcular métricas por equipe: {e}")
        metricas_equipe = pd.DataFrame()
    
    # Métricas de produtividade por usuário
    try:
        metricas_usuario = calcular_metricas_produtividade_usuario(
            dados,
            absences_df,
            inicio_mes,
            fim_mes
        )
    except Exception as e:
        st.error(f"Erro ao calcular métricas por usuário: {e}")
        metricas_usuario = pd.DataFrame()
    
    # Análise de risco de projetos
    risco_projetos = []
    for _, projeto in projects_df.iterrows():
        # Filtrar entradas de timesheet para o projeto
        entries_projeto = dados[dados['project_id'] == projeto['project_id']]
        
        # Calcular custo realizado
        custo_realizado = 0
        horas_projeto = 0
        for _, entry in entries_projeto.iterrows():
            try:
                user_rate_id = users_df[users_df['user_id'] == entry['user_id']]['rate_id'].iloc[0]
                user_rate = rates_df[rates_df['rate_id'] == user_rate_id]
                if not user_rate.empty:
                    taxa = float(user_rate['rate_cost'].iloc[0])
                    custo_entrada = entry['hours'] * taxa
                    custo_realizado += custo_entrada
                    horas_projeto += entry['hours']
            except Exception as e:
                st.warning(f"Erro ao calcular custo para projeto {projeto['project_name']}: {e}")
        
        # Calcular custo planejado proporcional
        data_inicio = pd.to_datetime(projeto['start_date'])
        data_fim = pd.to_datetime(projeto['end_date'])
        dias_projeto_total = calcular_dias_uteis_projeto(data_inicio.date(), data_fim.date())
        dias_projeto_ate_agora = calcular_dias_uteis_projeto(data_inicio.date(), fim_mes.date())
        
        custo_total_planejado = float(projeto.get('total_cost', 0))  # Usar get() para evitar erros
        
        # Adicionar log de debug
        st.sidebar.write(f"Projeto: {projeto['project_name']}")
        st.sidebar.write(f"Custo Total Planejado: €{custo_total_planejado:,.2f}")
        st.sidebar.write(f"Custo Realizado: €{custo_realizado:,.2f}")
        st.sidebar.write(f"Horas do Projeto: {horas_projeto:.2f}")
        st.sidebar.write(f"Dias do Projeto Total: {dias_projeto_total}")
        st.sidebar.write(f"Dias do Projeto até Agora: {dias_projeto_ate_agora}")
        
        # Calcular custo planejado proporcional
        custo_planejado_ate_agora = custo_total_planejado * (dias_projeto_ate_agora / dias_projeto_total) if dias_projeto_total > 0 else 0
        
        # Calcular risco
        risco = calcular_risco_projeto_dashboard(custo_realizado, custo_planejado_ate_agora)
        
        risco_projetos.append({
            'nome_projeto': projeto['project_name'],
            'nivel_risco': risco['nivel'],
            'cor_risco': risco['cor'],
            'custo_planejado': custo_planejado_ate_agora,
            'custo_realizado': custo_realizado,
            'variacao_percentual': risco['variacao_percentual']
        })
    
    risco_projetos_df = pd.DataFrame(risco_projetos)
    
    return {
        'metricas_equipe': metricas_equipe,
        'metricas_usuario': metricas_usuario,
        'risco_projetos': risco_projetos_df,
        'dias_uteis': dias_uteis
    }

# [resto do código permanece o mesmo]

def productivity_dashboard():
    """
    Dashboard consolidado de produtividade
    """
    st.title("📊 Painel Consolidado de Produtividade")
    
    # Filtros de período
    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox(
            "Mês",
            range(1, 13),
            format_func=lambda x: calendar.month_name[x],
            index=datetime.now().month - 1
        )
    with col2:
        ano = st.selectbox(
            "Ano",
            range(datetime.now().year - 2, datetime.now().year + 1),
            index=0
        )
    
    # Calcular indicadores
    indicadores = calcular_indicadores_consolidados(ano, mes)
    
    # Seção de Resumo Executivo
    st.header("Resumo Executivo")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            "Dias Úteis", 
            indicadores['dias_uteis'], 
            help="Dias úteis considerando feriados e fins de semana"
        )
    
    with col2:
        # Usar uma abordagem mais flexível para encontrar as horas realizadas
        if not indicadores['metricas_usuario'].empty:
            # Tenta diferentes nomes de colunas
            coluna_horas = [
                col for col in ['total_horas', 'hours', 'horas_realizadas'] 
                if col in indicadores['metricas_usuario'].columns
            ][0] if [col for col in ['total_horas', 'hours', 'horas_realizadas'] 
                      if col in indicadores['metricas_usuario'].columns] else None
            
            if coluna_horas:
                total_horas_realizadas = indicadores['metricas_usuario'][coluna_horas].sum()
                st.metric(
                    "Total de Horas Realizadas", 
                    f"{total_horas_realizadas:.2f}h",
                    help="Total de horas trabalhadas no período"
                )
            else:
                st.metric(
                    "Total de Horas Realizadas", 
                    "N/D",
                    help="Não foi possível calcular total de horas"
                )
        else:
            st.metric(
                "Total de Horas Realizadas", 
                "N/D",
                help="Sem dados de horas disponíveis"
            )
    
    with col3:
        if not indicadores['metricas_usuario'].empty:
            # Tenta diferentes nomes de colunas
            coluna_ocupacao = [
                col for col in ['percentual_ocupacao', 'ocupacao', 'perc_ocupacao'] 
                if col in indicadores['metricas_usuario'].columns
            ][0] if [col for col in ['percentual_ocupacao', 'ocupacao', 'perc_ocupacao'] 
                      if col in indicadores['metricas_usuario'].columns] else None
            
            if coluna_ocupacao:
                media_ocupacao = indicadores['metricas_usuario'][coluna_ocupacao].mean()
                st.metric(
                    "Ocupação Média", 
                    f"{media_ocupacao:.1f}%", 
                    help="Percentual médio de ocupação dos colaboradores"
                )
            else:
                st.metric(
                    "Ocupação Média", 
                    "N/D",
                    help="Não foi possível calcular ocupação média"
                )
        else:
            st.metric(
                "Ocupação Média", 
                "N/D",
                help="Sem dados de ocupação disponíveis"
            )
    
    # Seção de Produtividade por Equipe
    if not indicadores['metricas_equipe'].empty:
        st.header("Produtividade por Equipe")
        
        # Gráfico de Produtividade por Equipe
        fig_equipe = go.Figure()
        
        # Barras de Ocupação
        fig_equipe.add_trace(go.Bar(
            x=indicadores['metricas_equipe']['group_name'],
            y=indicadores['metricas_equipe']['percentual_ocupacao'],
            name='% Ocupação',
            marker_color='#8884d8'
        ))
        
        # Barras de Faturável
        fig_equipe.add_trace(go.Bar(
            x=indicadores['metricas_equipe']['group_name'],
            y=indicadores['metricas_equipe']['percentual_faturavel'],
            name='% Faturável',
            marker_color='#82ca9d'
        ))
        
        # Linha de Ausências
        fig_equipe.add_trace(go.Scatter(
            x=indicadores['metricas_equipe']['group_name'],
            y=indicadores['metricas_equipe']['percentual_ausencias'],
            name='% Ausências',
            mode='lines+markers',
            line=dict(color='#ff7f7f', width=2)
        ))
        
        # Linhas de meta
        fig_equipe.add_trace(go.Scatter(
            x=indicadores['metricas_equipe']['group_name'],
            y=[87.5] * len(indicadores['metricas_equipe']),
            name='Meta Ocupação (87.5%)',
            mode='lines',
            line=dict(color='red', width=2, dash='dash')
        ))
        
        fig_equipe.add_trace(go.Scatter(
            x=indicadores['metricas_equipe']['group_name'],
            y=[75] * len(indicadores['metricas_equipe']),
            name='Meta Faturável (75%)',
            mode='lines',
            line=dict(color='orange', width=2, dash='dash')
        ))
        
        fig_equipe.update_layout(
            title='Produtividade por Equipe',
            xaxis_title='Equipe',
            yaxis_title='Percentual (%)',
            barmode='group',
            height=500
        )
        
        st.plotly_chart(fig_equipe)
    
    # Seção de Risco de Projetos
    if not indicadores['risco_projetos'].empty:
        st.header("Risco de Projetos")
        
        # Gráfico de Risco de Projetos
        fig_risco = go.Figure(data=[
            go.Bar(
                x=indicadores['risco_projetos']['nome_projeto'],
                y=indicadores['risco_projetos']['variacao_percentual'],
                marker_color=indicadores['risco_projetos']['cor_risco'],
                text=indicadores['risco_projetos']['nivel_risco'],
                textposition='auto'
            )
        ])
        
        fig_risco.update_layout(
            title='Variação de Custo dos Projetos',
            xaxis_title='Projeto',
            yaxis_title='Variação de Custo (%)',
            height=500
        )
        
        st.plotly_chart(fig_risco)
        
        # Tabela de Risco de Projetos
        st.subheader("Detalhamento de Risco de Projetos")
        st.dataframe(
            indicadores['risco_projetos'].rename(columns={
                'nome_projeto': 'Projeto',
                'nivel_risco': 'Nível de Risco',
                'custo_planejado': 'Custo Planejado',
                'custo_realizado': 'Custo Realizado',
                'variacao_percentual': 'Variação (%)'
            }).style.format({
                'Custo Planejado': '€{:,.2f}',
                'Custo Realizado': '€{:,.2f}',
                'Variação (%)': '{:.2f}%'
            }).applymap(
                lambda x: 'background-color: #ffcccc' if x == 'Alto Risco' else 
                          'background-color: #fff2cc' if x == 'Risco Moderado' else '',
                subset=['Nível de Risco']
            )
        )
    else:
        st.warning("Não foi possível gerar o relatório de risco de projetos.")