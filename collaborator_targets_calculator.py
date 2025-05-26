import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
import plotly.express as px
import plotly.graph_objects as go
from database_manager import DatabaseManager
import io

def collaborator_targets_calculator_page():
    """
    Interface de cálculo de metas por colaborador com base nas metas anuais.
    Permite calcular metas individuais considerando rates, limite de horas 
    e metas definidas por empresa.
    """
    st.title("Cálculo de Metas por Colaborador")
    
    # Inicializar o gerenciador de banco de dados
    db_manager = DatabaseManager()
    
    # Constantes importantes
    MAX_BILLABLE_HOURS_PER_MONTH = 22 * 8 * 0.8  # ~140.8 horas/mês (22 dias úteis, 8h/dia, 80% faturáveis)
    EFFECTIVE_MONTHS = 11  # Considerando 1 mês de férias coletivas (agosto)

    # Criar tabs para diferentes funcionalidades
    tab_calcular, tab_visualizar, tab_exportar = st.tabs([
        "Calcular Metas", 
        "Visualizar Metas",
        "Exportar Relatório"
    ])
    
    with tab_calcular:
        st.header("Cálculo de Metas por Colaborador")
        
        # Seleção do ano
        col1, col2 = st.columns(2)
        with col1:
            ano = st.selectbox(
                "Ano de Referência",
                range(datetime.now().year, datetime.now().year + 3),
                index=0
            )
        
        # Verificar se já existem metas anuais para o ano selecionado
        annual_targets_df = db_manager.query_to_df(
            "SELECT * FROM annual_targets WHERE target_year = ?", 
            (ano,)
        )
        
        if annual_targets_df.empty:
            st.warning(f"Não foram encontradas metas anuais definidas para o ano {ano}.")
            
            # Botão para carregar a interface de definição de metas
            if st.button("Definir Metas Anuais"):
                # Redirecionamento para a página de metas anuais
                from annual_targets import annual_targets_page
                annual_targets_page()
            return
            
        # Carregar usuários ativos apenas dos grupos Tech, DS e LRB
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores WHERE active = 1")
        
        # Filtrar usuários dos grupos específicos
        tech_users = []
        ds_users = []
        lrb_users = []
        
        for _, user in users_df.iterrows():
            try:
                if not user['groups']:
                    continue
                    
                # Processar o campo groups
                groups = eval(str(user['groups'])) if isinstance(user['groups'], str) else user['groups']
                
                # Converter para lista se for dicionário ou outro tipo
                if isinstance(groups, dict):
                    groups = list(groups.values())
                elif not isinstance(groups, list):
                    groups = [groups]
                
                # Converter strings numéricas para inteiros
                groups = [int(g) if isinstance(g, str) and g.isdigit() else g for g in groups]
                
                # Verificar grupos em um formato que contemple várias possibilidades
                groups_str = str(groups).lower()
                
                if (1 in groups or 'tech' in groups_str or 'commercial' in groups_str or 'perform' in groups_str):
                    tech_users.append(user)
                elif (3 in groups or 'ds' in groups_str):
                    ds_users.append(user)
                elif (4 in groups or 'lrb' in groups_str):
                    lrb_users.append(user)
            except Exception as e:
                st.error(f"Erro ao processar grupos do usuário {user['user_id']}: {e}")
                # Opcionalmente, adicionar a um grupo padrão
                # tech_users.append(user)
        
        # Resumo de usuários por grupo
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Tech", len(tech_users))
        with col2:
            st.metric("Design", len(ds_users))
        with col3:
            st.metric("LRB", len(lrb_users))
            
        # Obter rates
        rates_df = db_manager.query_to_df("SELECT * FROM rates")
        
        # Se não encontrar usuários nos grupos específicos
        if len(tech_users) + len(ds_users) + len(lrb_users) == 0:
            st.error("Não foram encontrados usuários ativos nos grupos Tech, Design ou LRB.")
            return
        
        # Calculando médias das rates por grupo
        def calculate_avg_rate(users_list):
            total_rate = 0
            count = 0
            user_rates = []
            
            for user in users_list:
                if pd.notna(user['rate_id']):
                    rate_info = rates_df[rates_df['rate_id'] == user['rate_id']]
                    if not rate_info.empty:
                        rate_value = float(rate_info['rate_cost'].iloc[0])
                        total_rate += rate_value
                        count += 1
                        user_rates.append({
                            'user_id': user['user_id'],
                            'name': f"{user['First_Name']} {user['Last_Name']}",
                            'rate': rate_value
                        })
            
            avg_rate = total_rate / count if count > 0 else 0
            return avg_rate, user_rates
        
        avg_rate_tech, tech_rates = calculate_avg_rate(tech_users)
        avg_rate_ds, ds_rates = calculate_avg_rate(ds_users)
        avg_rate_lrb, lrb_rates = calculate_avg_rate(lrb_users)
        
        # Mostrar rates médias
        st.subheader("Rates Médias por Grupo")
        cols = st.columns(3)
        with cols[0]:
            st.metric("Tech", f"€{avg_rate_tech:.2f}/h")
        with cols[1]:
            st.metric("Design", f"€{avg_rate_ds:.2f}/h")
        with cols[2]:
            st.metric("LRB", f"€{avg_rate_lrb:.2f}/h")
        
        # Obter metas anuais
        tech_target = annual_targets_df[annual_targets_df['company_name'] == 'Tech']['target_value'].iloc[0] if 'Tech' in annual_targets_df['company_name'].values else 0
        ds_target = annual_targets_df[annual_targets_df['company_name'] == 'Design']['target_value'].iloc[0] if 'Design' in annual_targets_df['company_name'].values else 0
        lrb_target = annual_targets_df[annual_targets_df['company_name'] == 'LRB']['target_value'].iloc[0] if 'LRB' in annual_targets_df['company_name'].values else 0
        
        # Mostrar metas anuais
        st.subheader("Metas Anuais")
        cols = st.columns(3)
        with cols[0]:
            st.metric("Tech", f"€{tech_target:,.2f}")
        with cols[1]:
            st.metric("Design", f"€{ds_target:,.2f}")
        with cols[2]:
            st.metric("LRB", f"€{lrb_target:,.2f}")
        
        # Metas já calculadas
        existing_targets = db_manager.query_to_df(
            "SELECT * FROM collaborator_targets WHERE year = ?",
            (ano,)
        )
        
        if not existing_targets.empty:
            st.info(f"Já existem metas calculadas para o ano {ano}. Calcular novamente irá sobrescrever as metas existentes.")
        
        # Botão para calcular
        if st.button("Calcular Metas por Colaborador", type="primary"):
            with st.spinner("Calculando metas..."):
                # Lista para armazenar resultados
                results = []
                
                # Função para calcular metas para um grupo específico
                def calculate_targets_for_group(users, company_name, annual_target):
                    if len(users) == 0 or annual_target <= 0:
                        return []
                    
                    # Meta mensal total da empresa = Meta anual / Número de meses efetivos
                    monthly_company_target = annual_target / EFFECTIVE_MONTHS
                    st.write(f"Meta mensal da empresa {company_name}: €{monthly_company_target:.2f}")
                    
                    # Selecionar a rate média apropriada para a empresa
                    if company_name == 'Tech':
                        company_avg_rate = avg_rate_tech
                    elif company_name == 'Design':
                        company_avg_rate = avg_rate_ds
                    elif company_name == 'LRB':
                        company_avg_rate = avg_rate_lrb
                    else:
                        company_avg_rate = 0  # Valor padrão
                    
                    # Lista para armazenar dados temporários dos usuários
                    user_data_list = []
                    
                    # Calcular a capacidade total da equipe (soma da capacidade de cada colaborador)
                    total_capacity = 0
                    
                    for user in users:
                        # Obter rate do usuário
                        rate_value = None
                        if pd.notna(user['rate_id']):
                            rate_info = rates_df[rates_df['rate_id'] == user['rate_id']]
                            if not rate_info.empty:
                                rate_value = float(rate_info['rate_cost'].iloc[0])
                        
                        # Usar rate específica ou média da empresa
                        rate = rate_value if rate_value else company_avg_rate
                        
                        # Capacidade de gerar receita (baseada no limite máximo de horas)
                        capacity = rate * MAX_BILLABLE_HOURS_PER_MONTH
                        total_capacity += capacity
                        
                        # Armazenar dados temporários
                        user_data_list.append({
                            'user': user,
                            'rate': rate,
                            'capacity': capacity
                        })
                    
                    # Resultados finais
                    group_results = []
                    total_assigned = 0  # Variável para tracking
                    
                    # Distribuir a meta mensal entre os colaboradores com base na capacidade
                    for user_data in user_data_list:
                        user = user_data['user']
                        rate = user_data['rate']
                        capacity = user_data['capacity']
                        
                        # Calcular a meta mensal individual proporcional à capacidade
                        proportion = capacity / total_capacity if total_capacity > 0 else 0
                        individual_monthly_target = monthly_company_target * proportion
                        
                        # Calcular horas necessárias
                        billable_hours = individual_monthly_target / rate if rate > 0 else 0
                        
                        # Limitar às horas máximas permitidas
                        if billable_hours > MAX_BILLABLE_HOURS_PER_MONTH:
                            billable_hours = MAX_BILLABLE_HOURS_PER_MONTH
                            individual_monthly_target = billable_hours * rate
                            
                        # Rastrear o total atribuído
                        total_assigned += individual_monthly_target
                        
                        # Salvar para cada mês (exceto agosto)
                        for month in range(1, 13):
                            if month == 8:  # Agosto (férias coletivas)
                                continue
                                
                            # Verificar se já existe e deletar
                            db_manager.execute_query(
                                "DELETE FROM collaborator_targets WHERE user_id = ? AND year = ? AND month = ?",
                                (user['user_id'], ano, month)
                            )
                            
                            # Inserir novo target
                            db_manager.execute_query(
                                """
                                INSERT INTO collaborator_targets 
                                (user_id, year, month, billable_hours_target, revenue_target, company_name, created_at, updated_at)
                                VALUES (?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
                                """,
                                (user['user_id'], ano, month, billable_hours, individual_monthly_target, company_name)
                            )
                        
                        # Adicionar aos resultados
                        group_results.append({
                            'user_id': user['user_id'],
                            'name': f"{user['First_Name']} {user['Last_Name']}",
                            'email': user['email'],
                            'company': company_name,
                            'rate': rate,
                            'monthly_target': individual_monthly_target,
                            'monthly_hours': billable_hours,
                            'annual_target': individual_monthly_target * EFFECTIVE_MONTHS,
                            'annual_hours': billable_hours * EFFECTIVE_MONTHS
                        })
                    
                    # Verificar se o total de metas atribuídas equivale à meta mensal da empresa
                    difference_percentage = abs(total_assigned - monthly_company_target) / monthly_company_target * 100
                    
                    # Se a diferença for significativa (mais de 1%), registrar um aviso
                    if difference_percentage > 1:
                        st.warning(f"Atenção ({company_name}): Há uma diferença de {difference_percentage:.2f}% entre a meta mensal total calculada (€{total_assigned:.2f}) e a meta mensal esperada da empresa (€{monthly_company_target:.2f}).")
                    
                    return group_results
                
                # Calcular para cada grupo
                results.extend(calculate_targets_for_group(tech_users, 'Tech', tech_target))
                results.extend(calculate_targets_for_group(ds_users, 'Design', ds_target))
                results.extend(calculate_targets_for_group(lrb_users, 'LRB', lrb_target))
                
                # Converter para DataFrame para exibição
                results_df = pd.DataFrame(results)
                
                if not results_df.empty:
                    # Calcular estatísticas por empresa
                    stats_by_company = {}
                    for company in ['Tech', 'Design', 'LRB']:
                        company_data = results_df[results_df['company'] == company]
                        if not company_data.empty:
                            total_monthly_target = company_data['monthly_target'].sum()
                            total_monthly_hours = company_data['monthly_hours'].sum()
                            avg_rate = company_data['rate'].mean()
                            
                            stats_by_company[company] = {
                                'users': len(company_data),
                                'total_monthly_target': total_monthly_target,
                                'total_monthly_hours': total_monthly_hours,
                                'avg_rate': avg_rate,
                                'target_per_user': total_monthly_target / len(company_data),
                                'hours_per_user': total_monthly_hours / len(company_data)
                            }
                    
                    # Exibir resumo
                    st.success(f"Metas calculadas com sucesso para {len(results_df)} colaboradores!")
                    
                    st.subheader("Resumo por Empresa")
                    stats_data = []
                    for company, stats in stats_by_company.items():
                        stats_data.append({
                            'Empresa': company,
                            'Colaboradores': stats['users'],
                            'Meta Mensal Total': stats['total_monthly_target'],
                            'Horas Mensais Totais': stats['total_monthly_hours'],
                            'Rate Média': stats['avg_rate'],
                            'Meta Mensal por Colaborador': stats['target_per_user'],
                            'Horas Mensais por Colaborador': stats['hours_per_user']
                        })
                    
                    stats_df = pd.DataFrame(stats_data)
                    
                    # Estilizar DataFrame para melhor visualização
                    st.dataframe(
                        stats_df.style.format({
                            'Meta Mensal Total': '€{:.2f}',
                            'Horas Mensais Totais': '{:.2f}h',
                            'Rate Média': '€{:.2f}/h',
                            'Meta Mensal por Colaborador': '€{:.2f}',
                            'Horas Mensais por Colaborador': '{:.2f}h'
                        })
                    )
                    
                    # Gráfico de metas por empresa
                    fig = px.bar(
                        stats_df,
                        x='Empresa',
                        y='Meta Mensal Total',
                        title='Meta Mensal por Empresa',
                        color='Empresa',
                        text_auto='.2s'
                    )
                    fig.update_traces(texttemplate='€%{y:.2f}', textposition='outside')
                    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Guardar resultados na session state para uso nas outras tabs
                    st.session_state['collaborator_results'] = results_df
                    st.session_state['stats_by_company'] = stats_by_company
                else:
                    st.error("Não foi possível calcular as metas. Verifique se existem colaboradores e metas definidas.")
    
    with tab_visualizar:
        st.header("Visualização de Metas por Colaborador")
        
        # Verificar se temos resultados calculados (da sessão ou do banco)
        if 'collaborator_results' in st.session_state:
            results_df = st.session_state['collaborator_results']
        else:
            # Tentar obter do banco de dados
            collaborator_targets = db_manager.query_to_df(
                """
                SELECT ct.*, u.First_Name, u.Last_Name, u.email, r.rate_cost 
                FROM collaborator_targets ct
                JOIN utilizadores u ON ct.user_id = u.user_id
                LEFT JOIN rates r ON u.rate_id = r.rate_id
                WHERE ct.year = ?
                """,
                (ano,)
            )
            
            if collaborator_targets.empty:
                st.info("Não há dados de metas por colaborador. Por favor, calcule as metas na primeira tab.")
                return
            
            # Preparar dados similares ao resultado do cálculo
            results = []
            for user_id in collaborator_targets['user_id'].unique():
                user_data = collaborator_targets[collaborator_targets['user_id'] == user_id]
                if user_data.empty:
                    continue
                
                # Pegar primeira entrada para informações básicas
                first_entry = user_data.iloc[0]
                
                # Calcular médias/somas
                monthly_hours = user_data['billable_hours_target'].mean()
                monthly_target = user_data['revenue_target'].mean()
                
                results.append({
                    'user_id': user_id,
                    'name': f"{first_entry['First_Name']} {first_entry['Last_Name']}",
                    'email': first_entry['email'],
                    'company': first_entry['company_name'],
                    'rate': first_entry['rate_cost'],
                    'monthly_target': monthly_target,
                    'monthly_hours': monthly_hours,
                    'annual_target': monthly_target * EFFECTIVE_MONTHS,
                    'annual_hours': monthly_hours * EFFECTIVE_MONTHS
                })
            
            results_df = pd.DataFrame(results)
            
            if results_df.empty:
                st.info("Não foi possível recuperar dados de metas por colaborador.")
                return
        
        # Filtros
        filter_container = st.container()
        col1, col2 = filter_container.columns(2)
        
        with col1:
            company_filter = st.multiselect(
                "Filtrar por Empresa",
                options=['Tech', 'Design', 'LRB'],
                default=['Tech', 'Design', 'LRB']
            )
        
        with col2:
            sort_by = st.selectbox(
                "Ordenar por",
                options=["Colaborador", "Empresa", "Rate", "Meta Mensal", "Horas Mensais"],
                index=3
            )
        
        # Aplicar filtros
        filtered_df = results_df.copy()
        if company_filter:
            filtered_df = filtered_df[filtered_df['company'].isin(company_filter)]
        
        # Aplicar ordenação
        sort_columns = {
            "Colaborador": "name",
            "Empresa": "company",
            "Rate": "rate",
            "Meta Mensal": "monthly_target",
            "Horas Mensais": "monthly_hours"
        }
        
        filtered_df = filtered_df.sort_values(sort_columns[sort_by], ascending=(sort_by != "Meta Mensal" and sort_by != "Horas Mensais"))
        
        # Verificar limite de horas
        filtered_df['at_max_hours'] = filtered_df['monthly_hours'] >= MAX_BILLABLE_HOURS_PER_MONTH
        
        # Exibir tabela detalhada
        st.subheader("Detalhamento por Colaborador")
        
        # Preparar DataFrame para exibição
        display_df = filtered_df.copy()
        display_df['rate'] = display_df['rate'].apply(lambda x: f"€{x:.2f}/h")
        display_df['monthly_target'] = display_df['monthly_target'].apply(lambda x: f"€{x:.2f}")
        display_df['monthly_hours'] = display_df.apply(
            lambda x: f"{x['monthly_hours']:.2f}h {'(MAX)' if x['at_max_hours'] else ''}",
            axis=1
        )
        display_df['annual_target'] = display_df['annual_target'].apply(lambda x: f"€{x:.2f}")
        display_df['annual_hours'] = display_df['annual_hours'].apply(lambda x: f"{x:.2f}h")
        
        # Remover colunas não necessárias para exibição
        display_df = display_df.drop(columns=['user_id', 'email', 'at_max_hours'])
        
        # Renomear colunas
        display_df = display_df.rename(columns={
            'name': 'Colaborador',
            'company': 'Empresa',
            'rate': 'Rate',
            'monthly_target': 'Meta Mensal',
            'monthly_hours': 'Horas Mensais',
            'annual_target': 'Meta Anual',
            'annual_hours': 'Horas Anuais'
        })
        
        st.dataframe(display_df, use_container_width=True)
        
        # Gráficos
        st.subheader("Gráficos")
        
        # Preparar dados para gráficos
        chart_df = filtered_df.copy()
        chart_df['name'] = chart_df['name'].apply(lambda x: x if len(x) < 20 else x[:17] + '...')
        
        # Gráfico de metas mensais
        fig_target = px.bar(
            chart_df,
            x='name',
            y='monthly_target',
            color='company',
            title='Meta Mensal por Colaborador',
            labels={'name': 'Colaborador', 'monthly_target': 'Meta Mensal (€)', 'company': 'Empresa'},
            hover_data=['rate', 'monthly_hours', 'at_max_hours']
        )
        
        fig_target.update_layout(
            xaxis_tickangle=-45,
            yaxis_title='Meta Mensal (€)',
            xaxis_title='Colaborador',
            legend_title='Empresa',
            height=600
        )
        
        st.plotly_chart(fig_target, use_container_width=True)
        
        # Gráfico de horas mensais
        fig_hours = px.bar(
            chart_df,
            x='name',
            y='monthly_hours',
            color='company',
            title='Horas Mensais por Colaborador',
            labels={'name': 'Colaborador', 'monthly_hours': 'Horas Mensais', 'company': 'Empresa'},
            hover_data=['rate', 'monthly_target', 'at_max_hours']
        )
        
        # Adicionar linha para o máximo de horas
        fig_hours.add_shape(
            type="line",
            x0=-0.5,
            y0=MAX_BILLABLE_HOURS_PER_MONTH,
            x1=len(chart_df) - 0.5,
            y1=MAX_BILLABLE_HOURS_PER_MONTH,
            line=dict(
                color="Red",
                width=2,
                dash="dash",
            )
        )
        
        fig_hours.add_annotation(
            x=len(chart_df) - 1,
            y=MAX_BILLABLE_HOURS_PER_MONTH,
            text=f"Máximo: {MAX_BILLABLE_HOURS_PER_MONTH:.2f}h",
            showarrow=False,
            yshift=10
        )
        
        fig_hours.update_layout(
            xaxis_tickangle=-45,
            yaxis_title='Horas Mensais',
            xaxis_title='Colaborador',
            legend_title='Empresa',
            height=600
        )
        
        st.plotly_chart(fig_hours, use_container_width=True)
    
    with tab_exportar:
        st.header("Exportar Relatório")
        
        # Verificar se temos resultados calculados
        if 'collaborator_results' not in st.session_state:
            collaborator_targets = db_manager.query_to_df(
                """
                SELECT ct.*, u.First_Name, u.Last_Name, u.email, r.rate_cost 
                FROM collaborator_targets ct
                JOIN utilizadores u ON ct.user_id = u.user_id
                LEFT JOIN rates r ON u.rate_id = r.rate_id
                WHERE ct.year = ?
                """,
                (ano,)
            )
            
            if collaborator_targets.empty:
                st.info("Não há dados de metas por colaborador para exportar. Por favor, calcule as metas na primeira tab.")
                return
        
        # Opções de exportação
        export_format = st.radio(
            "Formato de Exportação",
            options=["Excel", "CSV"],
            horizontal=True
        )
        
        if st.button("Gerar Relatório", type="primary"):
            with st.spinner("Gerando relatório..."):
                # Obter dados
                if 'collaborator_results' in st.session_state:
                    results_df = st.session_state['collaborator_results']
                    stats_by_company = st.session_state.get('stats_by_company', {})
                else:
                    # Recuperar do banco de dados (semelhante à tab de visualização)
                    # Código similar à tab de visualização para recuperar dados
                    pass
                
                # Preparar dataframes para exportação
                export_df = results_df.copy()
                
                # Renomear colunas para relatório
                export_df = export_df.rename(columns={
                    'name': 'Colaborador',
                    'company': 'Empresa',
                    'rate': 'Rate (€/h)',
                    'monthly_target': 'Meta Mensal (€)',
                    'monthly_hours': 'Horas Mensais',
                    'annual_target': 'Meta Anual (€)',
                    'annual_hours': 'Horas Anuais',
                    'email': 'Email'
                })
                
                # Formatar valores numéricos
                for col in ['Meta Mensal (€)', 'Meta Anual (€)']:
                    export_df[col] = export_df[col].apply(lambda x: f"€{x:.2f}")
                    
                for col in ['Horas Mensais', 'Horas Anuais']:
                    export_df[col] = export_df[col].apply(lambda x: f"{x:.2f}h")
                    
                export_df['Rate (€/h)'] = export_df['Rate (€/h)'].apply(lambda x: f"€{x:.2f}")
                
                # Estatísticas por empresa (se disponível)
                if stats_by_company:
                    stats_data = []
                    for company, stats in stats_by_company.items():
                        stats_data.append({
                            'Empresa': company,
                            'Colaboradores': stats['users'],
                            'Meta Mensal Total (€)': f"€{stats['total_monthly_target']:.2f}",
                            'Horas Mensais Totais': f"{stats['total_monthly_hours']:.2f}h",
                            'Rate Média (€/h)': f"€{stats['avg_rate']:.2f}",
                            'Meta Mensal por Colaborador (€)': f"€{stats['target_per_user']:.2f}",
                            'Horas Mensais por Colaborador': f"{stats['hours_per_user']:.2f}h"
                        })
                    
                    stats_export_df = pd.DataFrame(stats_data)
                
                # Exportar conforme formato selecionado
                if export_format == "Excel":
                    # Importar io para tratar o buffer
                    import io
                    
                    # Criar arquivo Excel na memória
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Informações gerais
                        info_df = pd.DataFrame([{
                            'Ano': ano,
                            'Data de Geração': datetime.now().strftime('%d/%m/%Y %H:%M'),
                            'Meses Efetivos': EFFECTIVE_MONTHS,
                            'Máximo de Horas Mensais': f"{MAX_BILLABLE_HOURS_PER_MONTH:.2f}h"
                        }])
                        info_df.to_excel(writer, sheet_name='Informações', index=False)
                        
                        # Resumo por empresa (se disponível)
                        if stats_by_company:
                            stats_export_df.to_excel(writer, sheet_name='Resumo por Empresa', index=False)
                        
                        # Detalhamento por colaborador
                        export_df.to_excel(writer, sheet_name='Metas por Colaborador', index=False)
                    
                    # Preparar para download
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Baixar Relatório Excel",
                        data=output,
                        file_name=f"metas_colaboradores_{ano}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Relatório Excel gerado com sucesso!")
                else:  # CSV
                    # Converter DataFrame para CSV
                    csv = export_df.to_csv(index=False)
                    
                    st.download_button(
                        label="📥 Baixar Relatório CSV",
                        data=csv,
                        file_name=f"metas_colaboradores_{ano}.csv",
                        mime="text/csv"
                    )
                    
                    st.success("Relatório CSV gerado com sucesso!")
                
                # Seção para ajuda e dicas
                with st.expander("Ajuda e Dicas"):
                    st.markdown("""
                    ### Guia do Relatório
                    
                    Este relatório contém as seguintes informações:
                    
                    - **Informações Gerais**: Ano, data de geração, meses efetivos e limite de horas mensais
                    - **Resumo por Empresa**: Visão consolidada das metas por empresa (Tech, Design, LRB)
                    - **Metas por Colaborador**: Detalhamento individual com rates, metas mensais e anuais
                    
                    ### Como utilizar
                    
                    - Para análise financeira, concentre-se nas colunas de Meta Mensal e Meta Anual
                    - Para planejamento de capacidade, utilize as colunas de Horas Mensais e Horas Anuais
                    - Colaboradores com indicação (MAX) estão no limite de horas faturáveis mensais
                    - A rate de cada colaborador afeta diretamente sua contribuição para a meta financeira
                    
                    ### Metodologia de Cálculo
                    
                    - As metas anuais são divididas pelo número de meses efetivos (11, excluindo agosto)
                    - O valor mensal é dividido pelo número de colaboradores em cada grupo
                    - A meta individual é ajustada pela rate específica de cada colaborador
                    - Há um limite máximo de {:.2f} horas faturáveis por mês
                    - Quando um colaborador atinge o limite máximo de horas, sua meta é ajustada conforme sua rate
                    """.format(MAX_BILLABLE_HOURS_PER_MONTH))
                    
                    st.info("Dica: Exporte o relatório em Excel para ter acesso a todas as informações em um formato adequado para análises adicionais.")
    
    