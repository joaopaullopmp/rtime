# dashboard_debug.py
# Versão de depuração do dashboard para identificar problemas com os dados

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
from database_manager import DatabaseManager

def dashboard_debug():
    """
    Versão de depuração do dashboard para identificar problemas de dados
    """
    st.title("Depuração de Dados de Colaboradores")
    
    # Inicializar gerenciador de banco de dados
    db_manager = DatabaseManager()
    
    # Selecionar mês e ano de referência
    now = datetime.now()
    current_month = now.month
    current_year = now.year
    
    col1, col2 = st.columns(2)
    with col1:
        month = st.selectbox(
            "Mês",
            range(1, 13),
            format_func=lambda x: calendar.month_name[x],
            index=current_month - 1,
            key="debug_month"
        )
    with col2:
        year = st.selectbox(
            "Ano",
            range(current_year - 1, current_year + 2),
            index=1,
            key="debug_year"
        )
        
    # Carregar dados das equipes
    groups_df = db_manager.query_to_df("SELECT * FROM groups WHERE active = 1")
    
    # Filtro de equipe
    teams = ["Todas"] + groups_df["group_name"].tolist()
    selected_team = st.selectbox(
        "Equipa",
        options=teams,
        index=0,
        key="debug_team"
    )
    
    # Carregar dados necessários
    timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
    users_df = db_manager.query_to_df("SELECT * FROM utilizadores WHERE active = 1")
    
    # Definir período
    primeiro_dia = datetime(year, month, 1)
    ultimo_dia = datetime(year, month, calendar.monthrange(year, month)[1], 23, 59, 59)
    
    st.write(f"Período analisado: {primeiro_dia.strftime('%d/%m/%Y')} a {ultimo_dia.strftime('%d/%m/%Y')}")
    
    # Calcular dias úteis no mês, excluindo fins de semana
    dias_uteis = 0
    data_atual = primeiro_dia
    while data_atual <= ultimo_dia:
        if data_atual.weekday() < 5:  # 0-4 são dias úteis (seg-sex)
            dias_uteis += 1
        data_atual += timedelta(days=1)
    
    st.write(f"Dias úteis no mês: {dias_uteis} (= {dias_uteis * 8} horas úteis)")
    
    # Filtrar usuários por equipe, se necessário
    filtered_users = users_df.copy()
    if selected_team != "Todas":
        user_ids = []
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
                    user_ids.append(user['user_id'])
            except Exception as e:
                st.write(f"⚠️ Erro ao processar grupos do usuário {user['First_Name']} {user['Last_Name']}: {str(e)}")
                st.write(f"Valor de groups: {user['groups']}")
        
        filtered_users = users_df[users_df['user_id'].isin(user_ids)]
    
    if filtered_users.empty:
        st.warning("Não foram encontrados colaboradores com os filtros selecionados.")
        return
    
    st.write(f"Encontrado(s) {len(filtered_users)} colaborador(es) para a equipe {selected_team}")
    
    # Verificar a estrutura da tabela de timesheet
    st.subheader("Estrutura da tabela de timesheet")
    st.write("Colunas disponíveis:")
    st.write(timesheet_df.columns.tolist())
    
    # Verificar os tipos de dados
    st.write("Tipos de dados:")
    st.write(timesheet_df.dtypes)
    
    # Informações sobre valores boolean na coluna 'billable'
    if 'billable' in timesheet_df.columns:
        st.write("Valores únicos em 'billable':")
        st.write(timesheet_df['billable'].unique())
    
    # Analisar entradas por colaborador
    st.subheader("Análise por colaborador")
    
    for _, user in filtered_users.iterrows():
        user_id = user['user_id']
        user_name = f"{user['First_Name']} {user['Last_Name']}"
        
        st.write(f"### {user_name} (ID: {user_id})")
        
        # Converter datas para datetime na tabela
        try:
            timesheet_df['start_date_dt'] = pd.to_datetime(timesheet_df['start_date'], format='mixed', errors='coerce')
        except Exception as e:
            st.write(f"⚠️ Erro ao converter datas: {str(e)}")
            # Mostrar algumas datas de exemplo
            if not timesheet_df.empty:
                st.write("Exemplos de datas na tabela:")
                st.write(timesheet_df['start_date'].head().tolist())
        
        # Tentar diferentes abordagens para filtrar o timesheet
        st.write("#### Tentativa 1: Filtro por user_id")
        user_entries = timesheet_df[timesheet_df['user_id'] == user_id]
        st.write(f"Entradas encontradas: {len(user_entries)}")
        
        st.write("#### Tentativa 2: Filtro por user_id e período")
        try:
            period_entries = user_entries[
                (user_entries['start_date_dt'] >= primeiro_dia) &
                (user_entries['start_date_dt'] <= ultimo_dia)
            ]
            st.write(f"Entradas no período: {len(period_entries)}")
            
            # Mostrar total de horas
            total_horas = period_entries['hours'].sum() if not period_entries.empty else 0
            st.write(f"Total de horas no período: {total_horas:.2f}")
            
            # Tentar filtrar horas faturáveis
            if 'billable' in period_entries.columns:
                try:
                    # Testar diferentes métodos para filtrar boolean
                    st.write("#### Tentativa de filtrar horas faturáveis")
                    
                    # Método 1: Boolean direto
                    billable_entries1 = period_entries[period_entries['billable'] == True]
                    st.write(f"Método 1 (billable == True): {len(billable_entries1)} entradas, {billable_entries1['hours'].sum():.2f} horas")
                    
                    # Método 2: Valor 1
                    billable_entries2 = period_entries[period_entries['billable'] == 1]
                    st.write(f"Método 2 (billable == 1): {len(billable_entries2)} entradas, {billable_entries2['hours'].sum():.2f} horas")
                    
                    # Método 3: Texto "True"
                    billable_entries3 = period_entries[period_entries['billable'] == "True"]
                    st.write(f"Método 3 (billable == 'True'): {len(billable_entries3)} entradas, {billable_entries3['hours'].sum():.2f} horas")
                    
                    # Método 4: Conversão para bool
                    period_entries['billable_bool'] = period_entries['billable'].astype(bool)
                    billable_entries4 = period_entries[period_entries['billable_bool'] == True]
                    st.write(f"Método 4 (conversão para bool): {len(billable_entries4)} entradas, {billable_entries4['hours'].sum():.2f} horas")
                    
                except Exception as e:
                    st.write(f"⚠️ Erro ao filtrar horas faturáveis: {str(e)}")
            
            # Calcular percentuais
            horas_uteis_totais = dias_uteis * 8  # 8 horas por dia útil
            occupation_percentage = (total_horas / horas_uteis_totais * 100) if horas_uteis_totais > 0 else 0
            st.write(f"Percentual de ocupação: {occupation_percentage:.1f}%")
            
            # Mostrar as primeiras entradas para verificação
            if not period_entries.empty:
                st.write("Primeiras entradas no período:")
                st.write(period_entries.head())
            
        except Exception as e:
            st.write(f"⚠️ Erro ao filtrar por período: {str(e)}")
    
    # Verificar dias úteis
    st.subheader("Verificação de dias úteis")
    dias_no_mes = []
    for dia in range(1, calendar.monthrange(year, month)[1] + 1):
        data = datetime(year, month, dia)
        dia_semana = data.weekday()
        dias_no_mes.append({
            'data': data.strftime('%d/%m/%Y'),
            'dia_semana': calendar.day_name[dia_semana],
            'dia_util': dia_semana < 5,
        })
    
    st.write(pd.DataFrame(dias_no_mes))