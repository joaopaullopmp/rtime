import sqlite3
import pandas as pd
import streamlit as st
from database_manager import DatabaseManager
import calendar
from datetime import date, datetime

def create_annual_targets_table():
    """Cria a tabela de metas anuais se não existir"""
    db_manager = DatabaseManager()
    
    # Criar a tabela de metas anuais
    query = """
    CREATE TABLE IF NOT EXISTS annual_targets (
        target_id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT NOT NULL,
        target_value REAL NOT NULL,
        target_year INTEGER NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(company_name, target_year)
    );
    """
    db_manager.execute_query(query)
    
    # Criar tabela para armazenar os colaboradores associados às metas
    query_collaborators = """
    CREATE TABLE IF NOT EXISTS annual_targets_collaborators (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        target_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(target_id) REFERENCES annual_targets(target_id),
        FOREIGN KEY(user_id) REFERENCES utilizadores(user_id),
        UNIQUE(target_id, user_id)
    );
    """
    db_manager.execute_query(query_collaborators)
    
    print("Tabelas 'annual_targets' e 'annual_targets_collaborators' criadas ou verificadas.")

def calculate_working_days_in_year(year):
    """Calcula o número de dias úteis em um ano, considerando feriados em Portugal"""
    # Feriados nacionais fixos em Portugal
    holidays = [
        date(year, 1, 1),   # Ano Novo
        date(year, 2, 20),  # Dia da Liberdade
        date(year, 4, 25),  # Dia da Liberdade
        date(year, 5, 1),   # Dia do Trabalhador
        date(year, 6, 10),  # Dia de Portugal
        date(year, 8, 15),  # Assunção de Nossa Senhora
        date(year, 10, 5),  # Implantação da República
        date(year, 11, 1),  # Dia de Todos os Santos
        date(year, 12, 1),  # Restauração da Independência
        date(year, 12, 8),  # Dia da Imaculada Conceição
        date(year, 12, 25)  # Natal
    ]
    
    # Calcular os dias úteis no ano
    working_days = 0
    for month in range(1, 13):
        num_days = calendar.monthrange(year, month)[1]
        for day in range(1, num_days + 1):
            current_date = date(year, month, day)
            # Verificar se não é fim de semana e não é feriado
            if current_date.weekday() < 5 and current_date not in holidays:
                working_days += 1
    
    return working_days

class AnnualTargetManager:
    def __init__(self):
        self.db = DatabaseManager()
        self.table_name = 'annual_targets'
        self.collab_table = 'annual_targets_collaborators'
        
        # Garantir que a tabela existe
        create_annual_targets_table()
    
    def create(self, data, collaborators=None):
        """Cria um novo registro de meta anual com colaboradores associados"""
        # Verificar dados obrigatórios
        if 'company_name' not in data or 'target_value' not in data or 'target_year' not in data:
            raise ValueError("Dados obrigatórios: company_name, target_value, target_year")
        
        # Adicionar timestamp atual
        data['created_at'] = data.get('created_at', pd.Timestamp.now().isoformat())
        data['updated_at'] = data.get('updated_at', pd.Timestamp.now().isoformat())
        
        # Verificar se já existe meta para esta empresa/ano
        existing = self.get_target(data['company_name'], data['target_year'])
        if not existing.empty:
            target_id = existing.iloc[0]['target_id']
            self.update(target_id, data)
        else:
            # Preparar a inserção
            columns = ', '.join(data.keys())
            placeholders = ', '.join(['?'] * len(data))
            
            # Executar query de inserção
            query = f"INSERT INTO {self.table_name} ({columns}) VALUES ({placeholders})"
            cursor = self.db.execute_query(query, tuple(data.values()))
            target_id = cursor.lastrowid
        
        # Associar colaboradores à meta (se fornecidos)
        if collaborators and target_id:
            # Primeiro remover associações existentes
            self.remove_all_collaborators(target_id)
            
            # Adicionar novos colaboradores
            for user_id in collaborators:
                self.add_collaborator(target_id, user_id)
        
        return target_id
    
    def read(self, target_id=None):
        """Lê registros de metas"""
        if target_id is None:
            return self.db.query_to_df(f"SELECT * FROM {self.table_name} ORDER BY target_year DESC, company_name")
        else:
            return self.db.query_to_df(f"SELECT * FROM {self.table_name} WHERE target_id = ?", (target_id,))
    
    def get_target(self, company_name, target_year):
        """Obtém meta para uma empresa/ano específico"""
        return self.db.query_to_df(
            f"SELECT * FROM {self.table_name} WHERE company_name = ? AND target_year = ?", 
            (company_name, target_year)
        )
    
    def update(self, target_id, data):
        """Atualiza uma meta existente"""
        # Remover campos que não devem ser atualizados
        if 'created_at' in data:
            del data['created_at']
        
        # Atualizar timestamp
        data['updated_at'] = pd.Timestamp.now().isoformat()
        
        # Preparar a atualização
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        params = list(data.values())
        params.append(target_id)
        
        # Executar query de atualização
        query = f"UPDATE {self.table_name} SET {set_clause} WHERE target_id = ?"
        self.db.execute_query(query, params)
        return True
    
    def delete(self, target_id):
        """Exclui uma meta"""
        # Primeiro remover todas as associações de colaboradores
        self.remove_all_collaborators(target_id)
        
        # Depois remover a meta
        query = f"DELETE FROM {self.table_name} WHERE target_id = ?"
        self.db.execute_query(query, (target_id,))
        return True
    
    def get_company_targets(self, company_name):
        """Obtém todas as metas de uma empresa"""
        return self.db.query_to_df(
            f"SELECT * FROM {self.table_name} WHERE company_name = ? ORDER BY target_year DESC",
            (company_name,)
        )
    
    def get_year_targets(self, target_year):
        """Obtém todas as metas de um ano"""
        return self.db.query_to_df(
            f"SELECT * FROM {self.table_name} WHERE target_year = ? ORDER BY company_name",
            (target_year,)
        )
    
    def add_collaborator(self, target_id, user_id):
        """Adiciona um colaborador à meta anual"""
        current_time = pd.Timestamp.now().isoformat()
        query = f"""
        INSERT OR IGNORE INTO {self.collab_table} (target_id, user_id, created_at)
        VALUES (?, ?, ?)
        """
        self.db.execute_query(query, (target_id, user_id, current_time))
        return True
    
    def remove_collaborator(self, target_id, user_id):
        """Remove um colaborador da meta anual"""
        query = f"DELETE FROM {self.collab_table} WHERE target_id = ? AND user_id = ?"
        self.db.execute_query(query, (target_id, user_id))
        return True
    
    def remove_all_collaborators(self, target_id):
        """Remove todos os colaboradores associados a uma meta"""
        query = f"DELETE FROM {self.collab_table} WHERE target_id = ?"
        self.db.execute_query(query, (target_id,))
        return True
    
    def get_target_collaborators(self, target_id):
        """Obtém todos os colaboradores associados a uma meta"""
        query = f"""
        SELECT u.user_id, u.First_Name, u.Last_Name, u.email, u.rate_id
        FROM {self.collab_table} tc
        JOIN utilizadores u ON tc.user_id = u.user_id
        WHERE tc.target_id = ?
        """
        return self.db.query_to_df(query, (target_id,))
    
    def calculate_target_metrics(self, target_id, occupation_percentage=75):
        """Calcula as métricas para a meta com base nos colaboradores associados
        
        Args:
            target_id: ID da meta anual
            occupation_percentage: Percentual de ocupação esperado (padrão: 75%)
        
        Returns:
            Dictionary com as métricas calculadas
        """
        # Obter informações da meta
        target_info = self.read(target_id)
        if target_info.empty:
            return None
        
        target_info = target_info.iloc[0]
        target_year = target_info['target_year']
        target_value = float(target_info['target_value'])
        
        # Obter colaboradores associados
        collaborators = self.get_target_collaborators(target_id)
        if collaborators.empty:
            return {
                'target_id': target_id,
                'target_year': target_year,
                'target_value': target_value,
                'num_collaborators': 0,
                'working_days': 0,
                'available_days': 0,
                'avg_daily_rate': 0,
                'avg_hourly_rate': 0,
                'required_hourly_rate': 0,
                'occupation_percentage': occupation_percentage,
                'collaborators_list': []
            }
        
        # Obter taxas dos colaboradores
        rates_df = self.db.query_to_df("SELECT * FROM rates")
        
        # Calcular dias úteis no ano
        working_days = calculate_working_days_in_year(target_year)
        
        # Dias de férias por colaborador (22 dias padrão + 6 dias adicionais da empresa)
        vacation_days = 22 + 6
        
        # Dias úteis disponíveis por colaborador (considerando férias)
        available_days = working_days - vacation_days
        
        # Horas úteis por dia
        hours_per_day = 8
        
        # Total de horas disponíveis por colaborador (considerando ocupação personalizada)
        available_hours_per_collab = available_days * hours_per_day * (occupation_percentage / 100)
        
        # Total de horas disponíveis para todos os colaboradores
        total_available_hours = available_hours_per_collab * len(collaborators)
        
        # Taxa horária necessária para atingir a meta
        required_hourly_rate = target_value / total_available_hours if total_available_hours > 0 else 0
        
        # Verificar taxas atuais dos colaboradores
        collaborators_list = []
        total_actual_rate = 0
        
        for _, collab in collaborators.iterrows():
            user_id = collab['user_id']
            name = f"{collab['First_Name']} {collab['Last_Name']}"
            rate_id = collab['rate_id']
            
            # Buscar a taxa atual do colaborador
            if pd.notna(rate_id) and not rates_df.empty:
                rate_info = rates_df[rates_df['rate_id'] == rate_id]
                if not rate_info.empty:
                    hourly_rate = float(rate_info['rate_cost'].iloc[0])
                else:
                    hourly_rate = 0
            else:
                hourly_rate = 0
            
            total_actual_rate += hourly_rate
            
            # Calcular métricas individuais
            individual_daily_revenue = hourly_rate * hours_per_day * (occupation_percentage / 100)
            individual_annual_revenue = individual_daily_revenue * available_days
            
            collaborators_list.append({
                'user_id': user_id,
                'name': name,
                'hourly_rate': hourly_rate,
                'daily_rate': hourly_rate * hours_per_day,
                'daily_revenue': individual_daily_revenue,
                'annual_revenue': individual_annual_revenue
            })
        
        # Calcular médias
        avg_hourly_rate = total_actual_rate / len(collaborators) if len(collaborators) > 0 else 0
        avg_daily_rate = avg_hourly_rate * hours_per_day
        
        # Calcular métricas gerais
        return {
            'target_id': target_id,
            'target_year': target_year,
            'target_value': target_value,
            'num_collaborators': len(collaborators),
            'working_days': working_days,
            'available_days': available_days,
            'available_hours_per_collab': available_hours_per_collab,
            'total_available_hours': total_available_hours,
            'avg_hourly_rate': avg_hourly_rate,
            'avg_daily_rate': avg_daily_rate,
            'required_hourly_rate': required_hourly_rate,
            'required_daily_rate': required_hourly_rate * hours_per_day,
            'occupation_percentage': occupation_percentage,
            'collaborators_list': collaborators_list
        }

def annual_targets_page():
    """Página de gerenciamento de metas anuais"""
    st.title("Metas Anuais")
    
    # Inicializar gerenciador
    target_manager = AnnualTargetManager()
    
    # Verificar se é um administrador
    if st.session_state.user_info['role'].lower() != 'admin':
        st.warning("Apenas administradores podem gerenciar metas anuais.")
        return
    
    # Interface de gerenciamento
    action = st.radio("Ação", ["Visualizar Metas", "Cadastrar Meta", "Editar Meta", "Excluir Meta"])
    
    if action == "Visualizar Metas":
        targets_df = target_manager.read()
        
        if targets_df.empty:
            st.info("Nenhuma meta anual cadastrada.")
        else:
            # Mostrar filtros
            col1, col2 = st.columns(2)
            with col1:
                company_filter = st.selectbox(
                    "Filtrar por Empresa", 
                    options=["Todas"] + sorted(targets_df['company_name'].unique().tolist())
                )
            with col2:
                year_filter = st.selectbox(
                    "Filtrar por Ano",
                    options=["Todos"] + sorted(targets_df['target_year'].unique().tolist(), reverse=True)
                )
            
            # Aplicar filtros
            filtered_df = targets_df.copy()
            if company_filter != "Todas":
                filtered_df = filtered_df[filtered_df['company_name'] == company_filter]
            if year_filter != "Todos":
                filtered_df = filtered_df[filtered_df['target_year'] == year_filter]
            
            # Exibir dados
            st.dataframe(
                filtered_df[['company_name', 'target_year', 'target_value']]
                .rename(columns={
                    'company_name': 'Empresa',
                    'target_year': 'Ano',
                    'target_value': 'Meta (€)'
                })
                .sort_values(['Ano', 'Empresa'], ascending=[False, True])
                .style.format({
                    'Meta (€)': '€{:,.2f}'
                })
            )
            
            # Gráfico de metas
            import plotly.express as px
            import plotly.graph_objects as go
            
            if not filtered_df.empty:
                # Adicionar dados de colaboradores para cada meta
                metrics_data = []
                
                for _, row in filtered_df.iterrows():
                    metrics = target_manager.calculate_target_metrics(row['target_id'])
                    if metrics:
                        metrics_data.append({
                            'company_name': row['company_name'],
                            'target_year': row['target_year'],
                            'target_value': row['target_value'],
                            'num_collaborators': metrics['num_collaborators'],
                            'avg_hourly_rate': metrics['avg_hourly_rate'],
                            'required_hourly_rate': metrics['required_hourly_rate']
                        })
                
                metrics_df = pd.DataFrame(metrics_data)
                
                # Gráfico de barras com metas e colaboradores
                fig = go.Figure()
                
                # Barras para o valor da meta
                fig.add_trace(go.Bar(
                    x=metrics_df['company_name'] if year_filter != "Todos" else metrics_df['target_year'].astype(str),
                    y=metrics_df['target_value'],
                    name='Meta Anual (€)',
                    marker_color='#1E88E5',
                    text=metrics_df['target_value'].apply(lambda x: f"€{x:,.0f}"),
                    textposition='auto'
                ))
                
                # Adicionar número de colaboradores como texto sobre as barras
                for i, row in metrics_df.iterrows():
                    fig.add_annotation(
                        x=row['company_name'] if year_filter != "Todos" else str(row['target_year']),
                        y=row['target_value'] + (max(metrics_df['target_value']) * 0.05),
                        text=f"{row['num_collaborators']} colaboradores",
                        showarrow=False,
                        font=dict(size=10)
                    )
                
                fig.update_layout(
                    title=f"Metas Anuais - {year_filter if year_filter != 'Todos' else 'Todas as Empresas'}",
                    xaxis_title='Empresa' if year_filter != "Todos" else 'Ano',
                    yaxis_title='Meta (€)',
                    yaxis_tickformat='€,.0f',
                    height=500
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Segundo gráfico para taxas horárias
                fig2 = go.Figure()
                
                fig2.add_trace(go.Bar(
                    x=metrics_df['company_name'] if year_filter != "Todos" else metrics_df['target_year'].astype(str),
                    y=metrics_df['avg_hourly_rate'],
                    name='Custo Médio Atual (€/h)',
                    marker_color='#8BC34A',
                    text=metrics_df['avg_hourly_rate'].apply(lambda x: f"€{x:.2f}"),
                    textposition='auto'
                ))
                
                fig2.add_trace(go.Bar(
                    x=metrics_df['company_name'] if year_filter != "Todos" else metrics_df['target_year'].astype(str),
                    y=metrics_df['required_hourly_rate'],
                    name='Taxa Necessária (€/h)',
                    marker_color='#FF5722',
                    text=metrics_df['required_hourly_rate'].apply(lambda x: f"€{x:.2f}"),
                    textposition='auto'
                ))
                
                fig2.update_layout(
                    title=f"Taxas Horárias - {year_filter if year_filter != 'Todos' else 'Todas as Empresas'}",
                    xaxis_title='Empresa' if year_filter != "Todos" else 'Ano',
                    yaxis_title='Taxa Horária (€)',
                    barmode='group',
                    height=500
                )
                
                st.plotly_chart(fig2, use_container_width=True)
            
            # Mostrar detalhes da meta ao clicar
            st.subheader("Detalhes da Meta")
            target_options = []
            for _, row in filtered_df.iterrows():
                target_options.append({
                    "id": row['target_id'],
                    "name": f"{row['company_name']} - {row['target_year']} - €{row['target_value']:,.2f}"
                })
            
            if target_options:
                selected_target = st.selectbox(
                    "Selecione uma meta para ver detalhes",
                    options=target_options,
                    format_func=lambda x: x['name']
                )
                
                if selected_target:
                    target_id = selected_target['id']
                    
                    # Adicionar campo para ajustar a ocupação na visualização
                    occupation_percentage = st.slider(
                        "Ajustar Percentual de Ocupação (%)",
                        min_value=50,
                        max_value=100,
                        value=75,
                        step=5,
                        key="occupation_view"
                    )
                    
                    # Calcular métricas com a ocupação ajustada
                    metrics = target_manager.calculate_target_metrics(
                        target_id, 
                        occupation_percentage=occupation_percentage
                    )
                    
                    if metrics:
                        # Mostrar informações gerais
                        st.markdown("### Informações Gerais")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Meta Anual", f"€{metrics['target_value']:,.2f}")
                            st.metric("Número de Colaboradores", metrics['num_collaborators'])
                            #st.metric("Dias Úteis no Ano", metrics['working_days'])
                            st.metric("Dias Úteis Disponíveis por Colaborador", metrics['available_days'], 
                                    help="Dias úteis menos dias de férias (22+6)")
                        with col2:
                            st.metric("Custo Médio Atual (Hora)", f"€{metrics['avg_hourly_rate']:.2f}")
                            #st.metric("Taxa Média Atual (Dia)", f"€{metrics['avg_daily_rate']:.2f}")
                            
                            # Verificar se as chaves existem antes de tentar acessá-las
                            required_hourly_rate = metrics.get('required_hourly_rate', 0)
                            #required_daily_rate = metrics.get('required_daily_rate', 0)
                            
                            # Calcular a taxa diária necessária se ela não estiver disponível
                            #if 'required_daily_rate' not in metrics and 'required_hourly_rate' in metrics:
                                #required_daily_rate = metrics['required_hourly_rate'] * 8  # assumindo 8 horas por dia
                            
                            st.metric("Taxa Necessária (Hora)", f"€{required_hourly_rate:.2f}")
                            #st.metric("Taxa Necessária (Hora)", f"€{required_hourly_rate:.2f}", 
                                    #delta=f"{required_hourly_rate-metrics['avg_hourly_rate']:.2f}",
                                    #delta_color="inverse")
                            #st.metric("Taxa Necessária (Dia)", f"€{required_daily_rate:.2f}", 
                                    #delta=f"{required_daily_rate-metrics['avg_daily_rate']:.2f}",
                                    #delta_color="inverse")
                        
                        # Mostrar cards com estatísticas importantes
                        st.markdown("### Análise da Capacidade")
                        cards = st.columns(3)
                        
                        # Obter valores com fallback para evitar KeyError
                        available_hours_per_collab = metrics.get('available_hours_per_collab', 0)
                        total_available_hours = metrics.get('total_available_hours', 0)
                        available_days = metrics.get('available_days', 0)
                        avg_hourly_rate = metrics.get('avg_hourly_rate', 0)
                        target_value = metrics.get('target_value', 0)
                        occupation_percentage = metrics.get('occupation_percentage', 75)
                        
                        with cards[0]:
                            st.markdown(f"""
                            <div style="padding:10px;border-radius:5px;background-color:#E3F2FD">
                                <h4 style="margin:0;color:#1976D2">Horas Disponíveis</h4>
                                <p style="font-size:24px;font-weight:bold;margin:10px 0">
                                    {available_hours_per_collab:.1f} horas/colaborador
                                </p>
                                <p style="margin:0">
                                    <strong>Total:</strong> {total_available_hours:.1f} horas
                                </p>
                                <p style="font-size:12px;color:#555;margin-top:5px">
                                    Considerando {available_days} dias úteis, 8h/dia e {occupation_percentage}% de ocupação
                                </p>
                            </div>
                            """, unsafe_allow_html=True)

                        # Card 2 - Custo Total Estimado
                        with cards[1]:
                            estimated_revenue = avg_hourly_rate * total_available_hours
                            percentage = (estimated_revenue / target_value) * 100 if target_value > 0 else 0
                            
                            st.markdown(f"""
                            <div style="padding:10px;border-radius:5px;background-color:#E8F5E9">
                                <h4 style="margin:0;color:#388E3C">Custo Total Estimado</h4>
                                <p style="font-size:24px;font-weight:bold;margin:10px 0">
                                    €{estimated_revenue:,.2f}
                                </p>
                                <p style="margin:0">
                                    <strong>{percentage:.1f}%</strong> da meta anual
                                </p>
                                <p style="font-size:12px;color:#555;margin-top:5px">
                                    Com a taxa média atual de €{avg_hourly_rate:.2f}/hora
                                </p>
                            </div>
                            """, unsafe_allow_html=True)

                        # Card 3 - Receita Total Planejada
                        with cards[2]:
                            planned_revenue = required_hourly_rate * total_available_hours
                            
                            st.markdown(f"""
                            <div style="padding:10px;border-radius:5px;background-color:#E3F2FD">
                                <h4 style="margin:0;color:#1976D2">Receita Total Planejada</h4>
                                <p style="font-size:24px;font-weight:bold;margin:10px 0">
                                    €{planned_revenue:,.2f}
                                </p>
                                <p style="margin:0">
                                    <strong>100%</strong> da meta anual
                                </p>
                                <p style="font-size:12px;color:#555;margin-top:5px">
                                    Com a taxa necessária de €{required_hourly_rate:.2f}/hora
                                </p>
                            </div>
                            """, unsafe_allow_html=True)

                        # Add information about actual revenue needed
                        st.markdown("### Receita Necessária Atual")
                        st.info(
                            f"""
                            Considerando as horas já realizadas pelos colaboradores, para atingir a meta anual de €{target_value:,.2f}:
                            
                            - Taxa horária necessária: €{required_hourly_rate:.2f}/hora
                            - Esta taxa deve ser aplicada a todas as horas faturáveis registradas para contribuir adequadamente com a meta.
                            """
                        )
                        
                            
    elif action == "Cadastrar Meta":
        with st.form("target_form"):
            company_name = st.selectbox(
                "Empresa*",
                options=["Tech", "Design", "LRB", "Boost"],
                help="Selecione a empresa"
            )
            
            target_year = st.number_input(
                "Ano*",
                min_value=2020,
                max_value=2030,
                value=2025,
                step=1,
                help="Ano da meta"
            )
            
            target_value = st.number_input(
                "Valor da Meta (€)*",
                min_value=0.0,
                value=0.0,
                step=10000.0,
                format="%.2f",
                help="Valor da meta anual em euros"
            )
            
            # Seleção de parâmetros adicionais
            occupation_percentage = st.slider(
                "Percentual de Ocupação (%)*",
                min_value=50,
                max_value=100,
                value=75,
                step=5,
                help="Percentual de ocupação esperado para os colaboradores (padrão: 75%)"
            )
            users_df = DatabaseManager().query_to_df("SELECT * FROM utilizadores WHERE active = 1")
            
            # Preparar opções de colaboradores filtrados por grupo (empresa)
            filtered_users = []
            for _, user in users_df.iterrows():
                try:
                    # Obter os grupos do usuário
                    if isinstance(user['groups'], str):
                        user_groups = eval(user['groups'])
                    else:
                        user_groups = user['groups'] if user['groups'] is not None else []
                    
                    # Converter para lista caso seja outro tipo
                    if not isinstance(user_groups, list):
                        if isinstance(user_groups, dict):
                            user_groups = list(user_groups.values())
                        else:
                            user_groups = [user_groups]
                    
                    # Verificar se o usuário pertence à empresa selecionada
                    company_match = False
                    for group in user_groups:
                        if isinstance(group, str) and group.lower() == company_name.lower():
                            company_match = True
                            break
                        elif isinstance(group, (int, float)) and str(group) == str(company_name):
                            company_match = True
                            break
                    
                    if company_match:
                        filtered_users.append(user)
                except Exception as e:
                    st.warning(f"Erro ao processar grupos do usuário {user['First_Name']} {user['Last_Name']}: {str(e)}")
                    continue
            
            # Criar opções para multiselect
            user_options = []
            for user in filtered_users:
                user_options.append({
                    "id": user['user_id'],
                    "name": f"{user['First_Name']} {user['Last_Name']}"
                })
            
            selected_users = st.multiselect(
                "Selecionar Colaboradores*",
                options=user_options,
                format_func=lambda x: x['name'],
                help="Selecione os colaboradores para esta meta"
            )
            
            if st.form_submit_button("Cadastrar"):
                if not company_name or target_year < 2020 or target_value <= 0:
                    st.error("Por favor, preencha todos os campos obrigatórios corretamente.")
                elif not selected_users:
                    st.error("Por favor, selecione pelo menos um colaborador.")
                else:
                    data = {
                        "company_name": company_name,
                        "target_year": int(target_year),
                        "target_value": float(target_value)
                    }
                    
                    # Converter lista de usuários selecionados para IDs
                    user_ids = [user['id'] for user in selected_users]
                    
                    # Verificamos se ja existe uma meta para esta empresa/ano
                    existing = target_manager.get_target(company_name, target_year)
                    if not existing.empty:
                        st.warning(f"Já existe uma meta para {company_name} no ano {target_year}. Deseja sobrescrever?")
                        if st.button("Sim, sobrescrever"):
                            target_manager.update(existing.iloc[0]['target_id'], data)
                            target_manager.remove_all_collaborators(existing.iloc[0]['target_id'])
                            for user_id in user_ids:
                                target_manager.add_collaborator(existing.iloc[0]['target_id'], user_id)
                            
                            # Mostrar métricas calculadas com a ocupação personalizada
                            st.success(f"Meta para {company_name} ({target_year}) atualizada com sucesso!")
                            
                            # Calcular e mostrar métricas
                            if existing.iloc[0]['target_id']:
                                metrics = target_manager.calculate_target_metrics(
                                    existing.iloc[0]['target_id'], 
                                    occupation_percentage=occupation_percentage
                                )
                                if metrics:
                                    required_hourly_rate = metrics.get('required_hourly_rate', 0)
                                    st.info(f"Com uma ocupação de {occupation_percentage}%, a taxa horária necessária é de €{required_hourly_rate:.2f}/hora.")
                    else:
                        target_id = target_manager.create(data, user_ids)
                        st.success(f"Meta para {company_name} ({target_year}) cadastrada com sucesso!")
                        
                        # Calcular e mostrar métricas
                        if target_id:
                            metrics = target_manager.calculate_target_metrics(
                                target_id, 
                                occupation_percentage=occupation_percentage
                            )
                            if metrics:
                                required_hourly_rate = metrics.get('required_hourly_rate', 0)
                                st.info(f"Com uma ocupação de {occupation_percentage}%, a taxa horária necessária é de €{required_hourly_rate:.2f}/hora.")
    
    elif action == "Editar Meta":
        targets_df = target_manager.read()
        
        if targets_df.empty:
            st.info("Nenhuma meta anual cadastrada para edição.")
        else:
            # Criar opções para seleção
            options = []
            for _, row in targets_df.iterrows():
                options.append({
                    "id": row['target_id'],
                    "name": f"{row['company_name']} - {row['target_year']} - €{row['target_value']:,.2f}"
                })
            
            selected = st.selectbox(
                "Selecione a meta para editar",
                options=options,
                format_func=lambda x: x['name']
            )
            
            if selected:
                target_id = selected['id']
                target_data = target_manager.read(target_id).iloc[0]
                
                # Buscar colaboradores associados
                current_collabs = target_manager.get_target_collaborators(target_id)
                current_collab_ids = current_collabs['user_id'].tolist() if not current_collabs.empty else []
                
                with st.form("edit_target_form"):
                    company_name = st.selectbox(
                        "Empresa*",
                        options=["Tech", "Design", "LRB", "Boost"],
                        index=["Tech", "Design", "LRB", "Boost"].index(target_data['company_name']) if target_data['company_name'] in ["Tech", "Design", "LRB", "Boost"] else 0
                    )
                    
                    target_year = st.number_input(
                        "Ano*",
                        min_value=2020,
                        max_value=2030,
                        value=int(target_data['target_year']),
                        step=1
                    )
                    
                    target_value = st.number_input(
                        "Valor da Meta (€)*",
                        min_value=0.0,
                        value=float(target_data['target_value']),
                        step=10000.0,
                        format="%.2f"
                    )
                    
                    # Seleção de parâmetros adicionais
                    occupation_percentage = st.slider(
                        "Percentual de Ocupação (%)*",
                        min_value=50,
                        max_value=100,
                        value=75,
                        step=5,
                        help="Percentual de ocupação esperado para os colaboradores (padrão: 75%)"
                    )
                    users_df = DatabaseManager().query_to_df("SELECT * FROM utilizadores WHERE active = 1")
                    
                    # Preparar opções de colaboradores filtrados por grupo (empresa)
                    filtered_users = []
                    for _, user in users_df.iterrows():
                        try:
                            # Obter os grupos do usuário
                            if isinstance(user['groups'], str):
                                user_groups = eval(user['groups'])
                            else:
                                user_groups = user['groups'] if user['groups'] is not None else []
                            
                            # Converter para lista caso seja outro tipo
                            if not isinstance(user_groups, list):
                                if isinstance(user_groups, dict):
                                    user_groups = list(user_groups.values())
                                else:
                                    user_groups = [user_groups]
                            
                            # Verificar se o usuário pertence à empresa selecionada
                            company_match = False
                            for group in user_groups:
                                if isinstance(group, str) and group.lower() == company_name.lower():
                                    company_match = True
                                    break
                                elif isinstance(group, (int, float)) and str(group) == str(company_name):
                                    company_match = True
                                    break
                            
                            if company_match:
                                filtered_users.append(user)
                        except Exception as e:
                            st.warning(f"Erro ao processar grupos do usuário {user['First_Name']} {user['Last_Name']}: {str(e)}")
                            continue
                    
                    # Criar opções para multiselect
                    user_options = []
                    for user in filtered_users:
                        user_options.append({
                            "id": user['user_id'],
                            "name": f"{user['First_Name']} {user['Last_Name']}"
                        })
                    
                    # Pré-selecionar colaboradores já associados
                    default_selected = []
                    for option in user_options:
                        if option['id'] in current_collab_ids:
                            default_selected.append(option)
                    
                    selected_users = st.multiselect(
                        "Selecionar Colaboradores*",
                        options=user_options,
                        default=default_selected,
                        format_func=lambda x: x['name'],
                        help="Selecione os colaboradores para esta meta"
                    )
                    
                    if st.form_submit_button("Atualizar"):
                        if not company_name or target_year < 2020 or target_value <= 0:
                            st.error("Por favor, preencha todos os campos obrigatórios corretamente.")
                        elif not selected_users:
                            st.error("Por favor, selecione pelo menos um colaborador.")
                        else:
                            data = {
                                "company_name": company_name,
                                "target_year": int(target_year),
                                "target_value": float(target_value)
                            }
                            
                            # Converter lista de usuários selecionados para IDs
                            user_ids = [user['id'] for user in selected_users]
                            
                            # Verificar conflito apenas se empresa ou ano mudou
                            if (company_name != target_data['company_name'] or 
                                target_year != target_data['target_year']):
                                existing = target_manager.get_target(company_name, target_year)
                                if not existing.empty and existing.iloc[0]['target_id'] != target_id:
                                    st.error(f"Já existe uma meta para {company_name} no ano {target_year}.")
                                    st.stop()
                            
                            target_manager.update(target_id, data)
                            target_manager.remove_all_collaborators(target_id)
                            for user_id in user_ids:
                                target_manager.add_collaborator(target_id, user_id)
                            st.success("Meta e colaboradores atualizados com sucesso!")
                            
                            # Calcular e mostrar métricas com a ocupação personalizada
                            metrics = target_manager.calculate_target_metrics(
                                target_id, 
                                occupation_percentage=occupation_percentage
                            )
                            if metrics:
                                required_hourly_rate = metrics.get('required_hourly_rate', 0)
                                st.info(f"Com uma ocupação de {occupation_percentage}%, a taxa horária necessária é de €{required_hourly_rate:.2f}/hora.")
    
    elif action == "Excluir Meta":
        targets_df = target_manager.read()
        
        if targets_df.empty:
            st.info("Nenhuma meta anual cadastrada para exclusão.")
        else:
            # Criar opções para seleção
            options = []
            for _, row in targets_df.iterrows():
                options.append({
                    "id": row['target_id'],
                    "name": f"{row['company_name']} - {row['target_year']} - €{row['target_value']:,.2f}"
                })
            
            selected = st.selectbox(
                "Selecione a meta para excluir",
                options=options,
                format_func=lambda x: x['name']
            )
            
            if selected:
                target_id = selected['id']
                st.warning(f"Tem certeza que deseja excluir a meta: {selected['name']}?")
                
                if st.button("Sim, excluir meta"):
                    target_manager.delete(target_id)
                    st.success("Meta excluída com sucesso!")

if __name__ == "__main__":
    # Código para testar a criação da tabela
    create_annual_targets_table()
    print("Tabela de metas anuais criada ou verificada.")