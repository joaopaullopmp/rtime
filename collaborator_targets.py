# collaborator_targets.py
import sqlite3
import pandas as pd
import streamlit as st
from database_manager import DatabaseManager
from annual_targets import AnnualTargetManager

def create_collaborator_targets_table():
    """Cria a tabela de indicadores por colaborador"""
    db_manager = DatabaseManager()
    
    query = """
    CREATE TABLE IF NOT EXISTS collaborator_targets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        billable_hours_target REAL NOT NULL,
        revenue_target REAL NOT NULL,
        company_name TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(user_id, year, month),
        FOREIGN KEY(user_id) REFERENCES utilizadores(user_id)
    );
    """
    db_manager.execute_query(query)
    print("Tabela 'collaborator_targets' criada ou verificada.")

class CollaboratorTargetCalculator:
    def __init__(self):
        self.db = DatabaseManager()
        self.annual_targets = AnnualTargetManager()
        
        # Garantir que a tabela existe
        create_collaborator_targets_table()
    
    def calculate_targets(self, year):
        """Calcula os alvos de horas e receita para todos os colaboradores"""
        # 1. Obter as metas anuais das empresas
        annual_targets = self.annual_targets.get_year_targets(year)
        if annual_targets.empty:
            return False, "Não há metas anuais definidas para o ano selecionado."
        
        # 2. Obter todos os usuários ativos
        users_df = self.db.query_to_df(
            "SELECT user_id, First_Name, Last_Name, email, rate_id, groups, role FROM utilizadores WHERE active = 1"
        )
        if users_df.empty:
            return False, "Não há colaboradores ativos no sistema."
        
        # 3. Obter as rates de cada colaborador
        rates_df = self.db.query_to_df("SELECT * FROM rates")
        if rates_df.empty:
            return False, "Não há rates definidas no sistema."
        
        # 4. Mapear os colaboradores para suas respectivas empresas
        # Com base nas regras: Comercial e Perform nas metas da Tech, outros em seus grupos
        tech_users = []
        ds_users = []
        lrb_users = []
        
        for _, user in users_df.iterrows():
            # Processar o campo groups que pode estar em diferentes formatos
            try:
                if isinstance(user['groups'], str):
                    groups = eval(user['groups'])
                else:
                    groups = user['groups']
                
                # Converter para lista se for dicionário ou outro tipo
                if isinstance(groups, dict):
                    groups = list(groups.values())
                elif not isinstance(groups, list):
                    groups = [groups]
                
                # Converter strings numéricas para inteiros
                groups = [int(g) if isinstance(g, str) and g.isdigit() else g for g in groups]
                
                # Mapear para as empresas
                if 1 in groups or 'Comercial' in groups or 'commercial' in str(groups).lower() or 'Perform' in groups or 'perform' in str(groups).lower():
                    tech_users.append(user)
                elif 3 in groups or 'DS' in groups or 'ds' in str(groups).lower():
                    ds_users.append(user)
                elif 4 in groups or 'LRB' in groups or 'lrb' in str(groups).lower():
                    lrb_users.append(user)
                else:
                    # Se não conseguir determinar, coloca na Tech como padrão
                    tech_users.append(user)
            except Exception as e:
                print(f"Erro ao processar grupos do usuário {user['user_id']}: {e}")
                # Colocar na Tech como padrão em caso de erro
                tech_users.append(user)
        
        # 5. Calcular a média de rate por empresa
        def calculate_avg_rate(users_list):
            total_rate = 0
            count = 0
            for user in users_list:
                if pd.notna(user['rate_id']):
                    rate_info = rates_df[rates_df['rate_id'] == user['rate_id']]
                    if not rate_info.empty:
                        total_rate += float(rate_info['rate_cost'].iloc[0])
                        count += 1
            return total_rate / count if count > 0 else 0
        
        avg_rate_tech = calculate_avg_rate(tech_users)
        avg_rate_ds = calculate_avg_rate(ds_users)
        avg_rate_lrb = calculate_avg_rate(lrb_users)
        
        # 6. Calcular as metas por colaborador
        tech_target = annual_targets[annual_targets['company_name'] == 'Tech']['target_value'].iloc[0] if 'Tech' in annual_targets['company_name'].values else 0
        ds_target = annual_targets[annual_targets['company_name'] == 'DS']['target_value'].iloc[0] if 'DS' in annual_targets['company_name'].values else 0
        lrb_target = annual_targets[annual_targets['company_name'] == 'LRB']['target_value'].iloc[0] if 'LRB' in annual_targets['company_name'].values else 0
        
        # Considerar meses úteis (excluindo férias coletivas, etc.)
        effective_months = 11  # Considerando 1 mês de férias em média
        
        # Definir limite máximo de horas faturáveis por mês por colaborador
        # Considerando uma média de 22 dias úteis por mês e 8h por dia, com 80% de faturabilidade
        MAX_BILLABLE_HOURS_PER_MONTH = 22 * 8 * 0.8  # ~140 horas/mês
        
        # 7. Calcular e salvar os indicadores de cada colaborador
        results = []
        
        # Para Tech
        if len(tech_users) > 0 and tech_target > 0:
            monthly_target_tech = tech_target / effective_months
            target_per_user_tech = monthly_target_tech / len(tech_users)
            
            for user in tech_users:
                # Obter rate do usuário
                rate_value = None
                if pd.notna(user['rate_id']):
                    rate_info = rates_df[rates_df['rate_id'] == user['rate_id']]
                    if not rate_info.empty:
                        rate_value = float(rate_info['rate_cost'].iloc[0])
                
                # Usar rate do usuário ou média da empresa
                rate = rate_value if rate_value else avg_rate_tech
                
                # Calcular horas faturáveis necessárias
                if rate > 0:
                    billable_hours = target_per_user_tech / rate
                    
                    # Verificar se excede o limite máximo
                    if billable_hours > MAX_BILLABLE_HOURS_PER_MONTH:
                        # Ajustar para o máximo e recalcular a meta de receita
                        billable_hours = MAX_BILLABLE_HOURS_PER_MONTH
                        adjusted_target = billable_hours * rate
                    else:
                        adjusted_target = target_per_user_tech
                else:
                    billable_hours = 0
                    adjusted_target = 0
                
                # Salvar resultados para cada mês
                for month in range(1, 13):
                    # Pular mês de férias coletivas (Agosto)
                    if month == 8:  # Agosto
                        continue
                        
                    self.save_target(
                        user['user_id'],
                        year,
                        month,
                        billable_hours,
                        adjusted_target,
                        'Tech'
                    )
                
                # Adicionar aos resultados
                results.append({
                    'user_id': user['user_id'],
                    'name': f"{user['First_Name']} {user['Last_Name']}",
                    'email': user['email'],
                    'company': 'Tech',
                    'rate': rate,
                    'monthly_target': adjusted_target,
                    'monthly_hours': billable_hours
                })
        
        # Para DS - aplique a mesma lógica de limite de horas
        if len(ds_users) > 0 and ds_target > 0:
            monthly_target_ds = ds_target / effective_months
            target_per_user_ds = monthly_target_ds / len(ds_users)
            
            for user in ds_users:
                # Obter rate do usuário
                rate_value = None
                if pd.notna(user['rate_id']):
                    rate_info = rates_df[rates_df['rate_id'] == user['rate_id']]
                    if not rate_info.empty:
                        rate_value = float(rate_info['rate_cost'].iloc[0])
                
                # Usar rate do usuário ou média da empresa
                rate = rate_value if rate_value else avg_rate_ds
                
                # Calcular horas faturáveis necessárias
                if rate > 0:
                    billable_hours = target_per_user_ds / rate
                    
                    # Verificar se excede o limite máximo
                    if billable_hours > MAX_BILLABLE_HOURS_PER_MONTH:
                        # Ajustar para o máximo e recalcular a meta de receita
                        billable_hours = MAX_BILLABLE_HOURS_PER_MONTH
                        adjusted_target = billable_hours * rate
                    else:
                        adjusted_target = target_per_user_ds
                else:
                    billable_hours = 0
                    adjusted_target = 0
                
                # Salvar resultados para cada mês
                for month in range(1, 13):
                    # Pular mês de férias coletivas (Agosto)
                    if month == 8:  # Agosto
                        continue
                        
                    self.save_target(
                        user['user_id'],
                        year,
                        month,
                        billable_hours,
                        adjusted_target,
                        'DS'
                    )
                
                # Adicionar aos resultados
                results.append({
                    'user_id': user['user_id'],
                    'name': f"{user['First_Name']} {user['Last_Name']}",
                    'email': user['email'],
                    'company': 'DS',
                    'rate': rate,
                    'monthly_target': adjusted_target,
                    'monthly_hours': billable_hours
                })
        
        # Para LRB - aplique a mesma lógica de limite de horas
        if len(lrb_users) > 0 and lrb_target > 0:
            monthly_target_lrb = lrb_target / effective_months
            target_per_user_lrb = monthly_target_lrb / len(lrb_users)
            
            for user in lrb_users:
                # Obter rate do usuário
                rate_value = None
                if pd.notna(user['rate_id']):
                    rate_info = rates_df[rates_df['rate_id'] == user['rate_id']]
                    if not rate_info.empty:
                        rate_value = float(rate_info['rate_cost'].iloc[0])
                
                # Usar rate do usuário ou média da empresa
                rate = rate_value if rate_value else avg_rate_lrb
                
                # Calcular horas faturáveis necessárias
                if rate > 0:
                    billable_hours = target_per_user_lrb / rate
                    
                    # Verificar se excede o limite máximo
                    if billable_hours > MAX_BILLABLE_HOURS_PER_MONTH:
                        # Ajustar para o máximo e recalcular a meta de receita
                        billable_hours = MAX_BILLABLE_HOURS_PER_MONTH
                        adjusted_target = billable_hours * rate
                    else:
                        adjusted_target = target_per_user_lrb
                else:
                    billable_hours = 0
                    adjusted_target = 0
                
                # Salvar resultados para cada mês
                for month in range(1, 13):
                    # Pular mês de férias coletivas (Agosto)
                    if month == 8:  # Agosto
                        continue
                        
                    self.save_target(
                        user['user_id'],
                        year,
                        month,
                        billable_hours,
                        adjusted_target,
                        'LRB'
                    )
                
                # Adicionar aos resultados
                results.append({
                    'user_id': user['user_id'],
                    'name': f"{user['First_Name']} {user['Last_Name']}",
                    'email': user['email'],
                    'company': 'LRB',
                    'rate': rate,
                    'monthly_target': adjusted_target,
                    'monthly_hours': billable_hours
                })
        
        return True, pd.DataFrame(results)
    
    def save_target(self, user_id, year, month, billable_hours, revenue, company):
        """Salva ou atualiza o target para um colaborador em um mês específico"""
        try:
            # Verificar se já existe
            existing = self.db.query_to_df(
                "SELECT id FROM collaborator_targets WHERE user_id = ? AND year = ? AND month = ?",
                (user_id, year, month)
            )
            
            if existing.empty:
                # Inserir novo
                query = """
                INSERT INTO collaborator_targets 
                (user_id, year, month, billable_hours_target, revenue_target, company_name, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
                """
                self.db.execute_query(query, (user_id, year, month, billable_hours, revenue, company))
            else:
                # Atualizar existente
                query = """
                UPDATE collaborator_targets 
                SET billable_hours_target = ?, revenue_target = ?, company_name = ?, updated_at = datetime('now')
                WHERE user_id = ? AND year = ? AND month = ?
                """
                self.db.execute_query(query, (billable_hours, revenue, company, user_id, year, month))
                
            return True
        except Exception as e:
            print(f"Erro ao salvar target: {e}")
            return False
    
    def get_user_targets(self, user_id, year=None):
        """Obtém os targets de um usuário específico"""
        if year:
            query = "SELECT * FROM collaborator_targets WHERE user_id = ? AND year = ? ORDER BY month"
            return self.db.query_to_df(query, (user_id, year))
        else:
            query = "SELECT * FROM collaborator_targets WHERE user_id = ? ORDER BY year DESC, month"
            return self.db.query_to_df(query, (user_id,))
    
    def get_company_targets(self, company, year):
        """Obtém os targets de todos os usuários de uma empresa"""
        query = "SELECT * FROM collaborator_targets WHERE company_name = ? AND year = ? ORDER BY month, user_id"
        return self.db.query_to_df(query, (company, year))
    
    def get_performance_vs_target(self, user_id, year, month=None):
        """Calcula o desempenho vs. target para um usuário"""
        # Obter targets
        if month:
            targets = self.db.query_to_df(
                "SELECT * FROM collaborator_targets WHERE user_id = ? AND year = ? AND month = ?",
                (user_id, year, month)
            )
        else:
            targets = self.db.query_to_df(
                "SELECT * FROM collaborator_targets WHERE user_id = ? AND year = ?",
                (user_id, year)
            )
        
        if targets.empty:
            return pd.DataFrame()
        
        # Obter timesheet entries
        timesheet_query = """
        SELECT t.*, u.rate_id
        FROM timesheet t
        JOIN utilizadores u ON t.user_id = u.user_id
        WHERE t.user_id = ? AND t.billable = 1
        """
        params = [user_id]
        
        if month:
            timesheet_query += " AND strftime('%Y', t.start_date) = ? AND strftime('%m', t.start_date) = ?"
            params.extend([str(year), str(month).zfill(2)])
        else:
            timesheet_query += " AND strftime('%Y', t.start_date) = ?"
            params.append(str(year))
        
        timesheet_df = self.db.query_to_df(timesheet_query, params)
        
        if timesheet_df.empty:
            # Retornar apenas com targets e zeros para realizados
            performance = targets.copy()
            performance['billable_hours_actual'] = 0
            performance['revenue_actual'] = 0
            performance['hours_completion'] = 0
            performance['revenue_completion'] = 0
            return performance
        
        # Calcular horas e receitas por mês
        timesheet_df['month'] = pd.to_datetime(timesheet_df['start_date']).dt.month
        
        # Obter rates para calcular receita
        rates_df = self.db.query_to_df("SELECT * FROM rates")
        
        # Função para calcular receita de um registro
        def calculate_revenue(row):
            if pd.isna(row['rate_id']):
                return 0
            
            rate_info = rates_df[rates_df['rate_id'] == row['rate_id']]
            if rate_info.empty:
                return 0
            
            return float(row['hours']) * float(rate_info['rate_cost'].iloc[0])
        
        # Calcular receita
        timesheet_df['revenue'] = timesheet_df.apply(calculate_revenue, axis=1)
        
        # Agrupar por mês
        monthly_data = timesheet_df.groupby('month').agg({
            'hours': 'sum',
            'revenue': 'sum'
        }).reset_index()
        
        # Mesclar com targets
        result = []
        for _, target in targets.iterrows():
            # Encontrar dados reais do mês
            month_data = monthly_data[monthly_data['month'] == target['month']] if not monthly_data.empty else None
            
            billable_hours_actual = float(month_data['hours'].iloc[0]) if month_data is not None and not month_data.empty else 0
            revenue_actual = float(month_data['revenue'].iloc[0]) if month_data is not None and not month_data.empty else 0
            
            # Calcular percentuais de conclusão
            hours_completion = (billable_hours_actual / float(target['billable_hours_target'])) * 100 if float(target['billable_hours_target']) > 0 else 0
            revenue_completion = (revenue_actual / float(target['revenue_target'])) * 100 if float(target['revenue_target']) > 0 else 0
            
            # Adicionar ao resultado
            result.append({
                'id': target['id'],
                'user_id': target['user_id'],
                'year': target['year'],
                'month': target['month'],
                'billable_hours_target': float(target['billable_hours_target']),
                'billable_hours_actual': billable_hours_actual,
                'revenue_target': float(target['revenue_target']),
                'revenue_actual': revenue_actual,
                'hours_completion': hours_completion,
                'revenue_completion': revenue_completion,
                'company_name': target['company_name']
            })
        
        return pd.DataFrame(result)

def show_targets_dashboard():
    """Interface para visualização e cálculo de indicadores"""
    st.title("Indicadores de Desempenho")
    
    calculator = CollaboratorTargetCalculator()
    
    # Verificar se é um administrador
    is_admin = st.session_state.user_info['role'].lower() == 'admin'
    
    if is_admin:
        # Interface de gerenciamento e visualização
        tabs = st.tabs(["Gerenciamento", "Visão Geral", "Meu Desempenho"])
        
        with tabs[0]:
            st.subheader("Cálculo de Indicadores por Colaborador")
            
            year = st.number_input(
                "Ano",
                min_value=2020,
                max_value=2030,
                value=2025,
                step=1,
                key="calc_year"
            )
            
            if st.button("Calcular Indicadores", key="calc_button"):
                progress = st.progress(0)
                status = st.empty()
                
                status.text("Calculando indicadores...")
                success, result = calculator.calculate_targets(year)
                progress.progress(100)
                
                if success:
                    st.success(f"Indicadores calculados com sucesso para {len(result)} colaboradores!")
                    
                    # Mostrar resumo por empresa
                    companies = result['company'].unique()
                    company_stats = []
                    
                    for company in companies:
                        company_data = result[result['company'] == company]
                        stats = {
                            'Empresa': company,
                            'Colaboradores': len(company_data),
                            'Meta Mensal Total': company_data['monthly_target'].sum(),
                            'Horas Faturaveis Mensais': company_data['monthly_hours'].sum(),
                            'Rate Médio': company_data['rate'].mean()
                        }
                        company_stats.append(stats)
                    
                    st.subheader("Resumo por Empresa")
                    stats_df = pd.DataFrame(company_stats)
                    st.dataframe(
                        stats_df.style.format({
                            'Meta Mensal Total': '€{:,.2f}',
                            'Horas Faturaveis Mensais': '{:,.2f}',
                            'Rate Médio': '€{:,.2f}'
                        })
                    )
                    
                    # Mostrar detalhes por colaborador
                    st.subheader("Detalhe por Colaborador")
                    st.dataframe(
                        result[['name', 'email', 'company', 'rate', 'monthly_target', 'monthly_hours']]
                        .rename(columns={
                            'name': 'Nome',
                            'email': 'Email',
                            'company': 'Empresa',
                            'rate': 'Rate/Hora',
                            'monthly_target': 'Meta Mensal (€)',
                            'monthly_hours': 'Horas Faturáveis/Mês'
                        })
                        .sort_values(['Empresa', 'Nome'])
                        .style.format({
                            'Rate/Hora': '€{:,.2f}',
                            'Meta Mensal (€)': '€{:,.2f}',
                            'Horas Faturáveis/Mês': '{:,.2f}'
                        })
                    )
                    
                    # Gráfico de horas faturáveis por empresa
                    import plotly.express as px
                    
                    fig_hours = px.bar(
                        result.groupby('company')['monthly_hours'].sum().reset_index(),
                        x='company',
                        y='monthly_hours',
                        labels={
                            'company': 'Empresa',
                            'monthly_hours': 'Horas Faturáveis Mensais'
                        },
                        title='Horas Faturáveis Mensais por Empresa',
                        color='company'
                    )
                    
                    st.plotly_chart(fig_hours, use_container_width=True)
                    
                    # Gráfico de distribuição de metas
                    fig_targets = px.pie(
                        result.groupby('company')['monthly_target'].sum().reset_index(),
                        values='monthly_target',
                        names='company',
                        title='Distribuição das Metas Mensais por Empresa'
                    )
                    
                    st.plotly_chart(fig_targets, use_container_width=True)
                else:
                    st.error(f"Erro ao calcular indicadores: {result}")
        
        with tabs[1]:
            st.subheader("Visão Geral de Desempenho")
            
            # Selecionar ano e empresa
            col1, col2 = st.columns(2)
            with col1:
                view_year = st.number_input(
                    "Ano",
                    min_value=2020,
                    max_value=2030,
                    value=2025,
                    step=1,
                    key="view_year"
                )
            with col2:
                view_company = st.selectbox(
                    "Empresa",
                    options=["Tech", "Design", "LRB", "Todas"],
                    key="view_company"
                )
            
            # Carregar dados de colaboradores
            users_df = calculator.db.query_to_df("SELECT user_id, First_Name, Last_Name FROM utilizadores WHERE active = 1")
            
            if users_df.empty:
                st.warning("Não há colaboradores ativos no sistema.")
            else:
                # Criar mapa de usuários para facilitar lookup
                users_map = {}
                for _, user in users_df.iterrows():
                    users_map[user['user_id']] = f"{user['First_Name']} {user['Last_Name']}"
                
                # Carregar dados de desempenho
                if view_company == "Todas":
                    # Todos os targets do ano
                    performance_data = calculator.db.query_to_df(
                        "SELECT * FROM collaborator_targets WHERE year = ?",
                        (view_year,)
                    )
                else:
                    # Targets de uma empresa específica
                    performance_data = calculator.db.query_to_df(
                        "SELECT * FROM collaborator_targets WHERE company_name = ? AND year = ?",
                        (view_company, view_year)
                    )
                
                if performance_data.empty:
                    st.warning(f"Não há dados de indicadores para {view_company} em {view_year}.")
                else:
                    # Processar para obter desempenho real
                    performance_results = []
                    
                    for user_id in performance_data['user_id'].unique():
                        user_performance = calculator.get_performance_vs_target(user_id, view_year)
                        performance_results.append(user_performance)
                    
                    if performance_results:
                        # Combinar resultados
                        all_performance = pd.concat(performance_results, ignore_index=True)
                        
                        # Verificar se temos dados
                        if all_performance.empty:
                            st.warning("Não foram encontrados dados de desempenho para o período selecionado.")
                            st.info("Tente selecionar outro período ou verifique se há registros no sistema.")
                        else:
                            # Verificar colunas disponíveis
                            required_columns = ['user_id', 'month', 'company_name', 
                                              'billable_hours_target', 'billable_hours_actual',
                                              'revenue_target', 'revenue_actual']

                            # Verificar quais colunas estão faltando
                            missing_columns = [col for col in required_columns if col not in all_performance.columns]

                            if missing_columns:
                                st.warning(f"Faltam as seguintes colunas nos dados: {missing_columns}")
                                st.info("Alguns dados podem não ser exibidos corretamente.")
                                
                                # Adicionar colunas ausentes com valores padrão
                                for col in missing_columns:
                                    if col == 'user_id':
                                        all_performance['user_id'] = 0
                                    elif col == 'month':
                                        all_performance['month'] = 1  # Janeiro
                                    elif col == 'company_name':
                                        all_performance['company_name'] = "Desconhecida"
                                    elif col in ['billable_hours_target', 'billable_hours_actual', 
                                               'revenue_target', 'revenue_actual']:
                                        all_performance[col] = 0.0
                            
                            # Adicionar nomes de usuários
                            all_performance['user_name'] = all_performance['user_id'].map(
                                lambda uid: users_map.get(uid, "Usuário Desconhecido")
                            )
                            
                            # Converter mês para nome
                            import calendar
                            all_performance['month_name'] = all_performance['month'].apply(
                                lambda x: calendar.month_name[x] if 1 <= x <= 12 else "Mês Desconhecido"
                            )
                            
                            # Resumo geral
                            st.subheader("Resumo de Desempenho")
                            
                            # Agrupar por empresa
                            try:
                                company_performance = all_performance.groupby('company_name').agg({
                                    'billable_hours_target': 'sum',
                                    'billable_hours_actual': 'sum',
                                    'revenue_target': 'sum',
                                    'revenue_actual': 'sum'
                                }).reset_index()
                                
                                # Calcular percentuais com verificação de divisão por zero
                                company_performance['hours_completion'] = company_performance.apply(
                                    lambda row: (row['billable_hours_actual'] / row['billable_hours_target'] * 100) 
                                               if row['billable_hours_target'] > 0 else 0,
                                    axis=1
                                )
                                
                                company_performance['revenue_completion'] = company_performance.apply(
                                    lambda row: (row['revenue_actual'] / row['revenue_target'] * 100) 
                                               if row['revenue_target'] > 0 else 0,
                                    axis=1
                                )
                                
                                # Mostrar resumo
                                st.dataframe(
                                    company_performance.rename(columns={
                                        'company_name': 'Empresa',
                                        'billable_hours_target': 'Meta de Horas',
                                        'billable_hours_actual': 'Horas Realizadas',
                                        'revenue_target': 'Meta de Receita',
                                        'revenue_actual': 'Receita Realizada',
                                        'hours_completion': '% Conclusão (Horas)',
                                        'revenue_completion': '% Conclusão (Receita)'
                                    }).style.format({
                                        'Meta de Horas': '{:,.2f}',
                                        'Horas Realizadas': '{:,.2f}',
                                        'Meta de Receita': '€{:,.2f}',
                                        'Receita Realizada': '€{:,.2f}',
                                        '% Conclusão (Horas)': '{:,.2f}%',
                                        '% Conclusão (Receita)': '{:,.2f}%'
                                    }).background_gradient(
                                        cmap='RdYlGn',
                                        subset=['% Conclusão (Horas)', '% Conclusão (Receita)'],
                                        vmin=0,
                                        vmax=100
                                    )
                                )
                                
                                # Gráfico de progresso
                                import plotly.express as px
                                
                                fig = px.bar(
                                    company_performance,
                                    x='company_name',
                                    y=['revenue_completion', 'hours_completion'],
                                    labels={
                                        'company_name': 'Empresa',
                                        'value': 'Percentual de Conclusão',
                                        'variable': 'Métrica'
                                    },
                                    title='Percentual de Conclusão por Empresa',
                                    barmode='group',
                                    color_discrete_map={
                                        'revenue_completion': '#1E88E5',
                                        'hours_completion': '#4CAF50'
                                    }
                                )
                                
                                # Adicionar linha de referência (100%)
                                fig.add_shape(
                                    type='line',
                                    x0=-0.5,
                                    x1=len(company_performance)-0.5,
                                    y0=100,
                                    y1=100,
                                    line=dict(color='red', width=2, dash='dash')
                                )
                                
                                st.plotly_chart(fig, use_container_width=True)
                                
                                # Top performers
                                st.subheader("Top Performers")
                                
                                # Agrupar por usuário
                                user_performance = all_performance.groupby(['user_id', 'user_name', 'company_name']).agg({
                                    'billable_hours_target': 'sum',
                                    'billable_hours_actual': 'sum',
                                    'revenue_target': 'sum',
                                    'revenue_actual': 'sum'
                                }).reset_index()
                                
                                # Calcular percentuais com verificação contra divisão por zero
                                user_performance['hours_completion'] = user_performance.apply(
                                    lambda row: (row['billable_hours_actual'] / row['billable_hours_target'] * 100) 
                                               if row['billable_hours_target'] > 0 else 0,
                                    axis=1
                                )
                                
                                user_performance['revenue_completion'] = user_performance.apply(
                                    lambda row: (row['revenue_actual'] / row['revenue_target'] * 100) 
                                               if row['revenue_target'] > 0 else 0,
                                    axis=1
                                )
                                
                                # Ordenar por percentual de receita
                                top_revenue = user_performance.sort_values('revenue_completion', ascending=False).head(10)
                                
                                st.dataframe(
                                    top_revenue[['user_name', 'company_name', 'revenue_target', 'revenue_actual', 'revenue_completion']]
                                    .rename(columns={
                                        'user_name': 'Colaborador',
                                        'company_name': 'Empresa',
                                        'revenue_target': 'Meta de Receita',
                                        'revenue_actual': 'Receita Realizada',
                                        'revenue_completion': '% Conclusão'
                                    })
                                    .style.format({
                                        'Meta de Receita': '€{:,.2f}',
                                        'Receita Realizada': '€{:,.2f}',
                                        '% Conclusão': '{:,.2f}%'
                                    })
                                    .background_gradient(
                                        cmap='RdYlGn',
                                        subset=['% Conclusão'],
                                        vmin=0,
                                        vmax=150
                                    )
                                )
                                
                                # Tendência mensal
                                st.subheader("Tendência Mensal")
                                
                                # Agrupar por mês
                                monthly_trend = all_performance.groupby(['month', 'month_name']).agg({
                                    'billable_hours_target': 'sum',
                                    'billable_hours_actual': 'sum',
                                    'revenue_target': 'sum',
                                    'revenue_actual': 'sum'
                                }).reset_index()
                                
                                # Calcular percentuais com proteção contra divisão por zero
                                monthly_trend['hours_completion'] = monthly_trend.apply(
                                    lambda row: (row['billable_hours_actual'] / row['billable_hours_target'] * 100) 
                                               if row['billable_hours_target'] > 0 else 0,
                                    axis=1
                                )
                                
                                monthly_trend['revenue_completion'] = monthly_trend.apply(
                                    lambda row: (row['revenue_actual'] / row['revenue_target'] * 100) 
                                               if row['revenue_target'] > 0 else 0,
                                    axis=1
                                )
                                
                                # Ordenar por mês
                                monthly_trend = monthly_trend.sort_values('month')
                                
                                # Gráfico de tendência
                                fig_trend = px.line(
                                    monthly_trend,
                                    x='month_name',
                                    y=['revenue_completion', 'hours_completion'],
                                    labels={
                                        'month_name': 'Mês',
                                        'value': 'Percentual de Conclusão',
                                        'variable': 'Métrica'
                                    },
                                    title='Tendência de Conclusão Mensal',
                                    markers=True
                                )
                                
                                # Adicionar linha de referência (100%)
                                fig_trend.add_shape(
                                    type='line',
                                    x0=0,
                                    x1=len(monthly_trend['month_name'])-1,
                                    y0=100,
                                    y1=100,
                                    line=dict(color='red', width=2, dash='dash')
                                )
                                
                                st.plotly_chart(fig_trend, use_container_width=True)
                                
                            except Exception as e:
                                st.error(f"Erro ao processar dados: {str(e)}")
                                import traceback
                                st.code(traceback.format_exc(), language="python")
                    else:
                        st.warning("Não foi possível calcular o desempenho. Verifique se há dados de timesheet registrados.")