import streamlit as st
import pandas as pd
import hashlib  # ← ADICIONAR ESTA LINHA
from datetime import datetime
from rate_manager import rate_page
from database_manager import UserManager, GroupManager, ClientManager, ProjectManager, TaskCategoryManager, RateManager

def unified_hash_password(password):
    """Hash unificado para todo o sistema"""
    return hashlib.sha256(str(password).encode()).hexdigest()

def get_team_rates():
    """Obter rates das equipes configurados no sistema"""
    try:
        rate_manager = RateManager()
        rates_df = rate_manager.read()
        
        team_rates = {}
        if not rates_df.empty:
            for _, rate in rates_df.iterrows():
                try:
                    team_name = str(rate.get('rate_name', '')).strip()
                    rate_cost = float(rate.get('rate_cost', 0))
                    if team_name and rate_cost > 0:
                        team_rates[team_name] = rate_cost
                except (ValueError, TypeError, AttributeError):
                    continue
        
        # Valores padrão caso não existam no banco ou em caso de erro
        if not team_rates:
            team_rates = {
                'Tech': 95.80,
                'LRB': 27.88,
                'Consultoria': 120.00,
                'Júnior': 45.00
            }
        
        return team_rates
    except Exception as e:
        print(f"Erro ao carregar rates das equipes: {e}")
        return {
            'Tech': 95.80,
            'LRB': 27.88,
            'Consultoria': 120.00,
            'Júnior': 45.00
        }

def render_project_form(is_edit=False, item=None, calculation_mode="Rate × Horas = Total", form_key=""):
    """Renderizar formulário de projeto (criar/editar)"""
    # Carregar dados necessários
    clients_manager = ClientManager()
    clients_df = clients_manager.read()
    team_rates = get_team_rates()
    
    # Valores padrão para criação ou valores atuais para edição
    if is_edit and item is not None:
        default_name = str(item.get("project_name", ""))
        default_description = str(item.get("project_description", ""))
        default_type = str(item.get("project_type", "Desenvolvimento"))
        default_status = str(item.get("status", "active"))
        
        # Conversão segura para float
        try:
            default_hourly_rate = float(item.get("hourly_rate", 0))
        except (ValueError, TypeError):
            default_hourly_rate = 0.0
            
        try:
            default_total_hours = float(item.get("total_hours", 0))
        except (ValueError, TypeError):
            default_total_hours = 0.0
            
        try:
            default_total_cost = float(item.get("total_cost", 0))
        except (ValueError, TypeError):
            default_total_cost = 0.0
            
        try:
            default_migrated_hours = float(item.get("horas_realizadas_mig", 0))
        except (ValueError, TypeError):
            default_migrated_hours = 0.0
            
        try:
            default_migrated_cost = float(item.get("custo_realizado_mig", 0))
        except (ValueError, TypeError):
            default_migrated_cost = 0.0
        
        # Encontrar cliente atual
        try:
            client_matches = clients_df[clients_df["client_id"] == item["client_id"]]
            default_client = client_matches["name"].iloc[0] if not client_matches.empty else ""
        except (KeyError, IndexError):
            default_client = ""
        
        # Converter datas
        try:
            if isinstance(item.get("start_date"), str):
                default_start_date = pd.to_datetime(item["start_date"]).date()
            else:
                default_start_date = item.get("start_date", datetime.now().date())
                
            if isinstance(item.get("end_date"), str):
                default_end_date = pd.to_datetime(item["end_date"]).date()
            else:
                default_end_date = item.get("end_date", datetime.now().date())
        except Exception as e:
            st.warning(f"Erro ao converter datas: {e}. Usando datas padrão.")
            default_start_date = datetime.now().date()
            default_end_date = datetime.now().date()
    else:
        default_name = ""
        default_description = ""
        default_type = "Desenvolvimento"
        default_status = "active"
        default_hourly_rate = 0.0
        default_total_hours = 0.0
        default_total_cost = 0.0
        default_migrated_hours = 0.0
        default_migrated_cost = 0.0
        default_client = ""
        default_start_date = datetime.now().date()
        default_end_date = datetime.now().date()
    
    # Lista de tipos de projeto centralizada
    project_types = ["Desenvolvimento", "Manutenção", "Consultoria", "OutTasking", "Bolsa Horas", "Projeto", "Interno"]
    
    # Campos básicos
    name = st.text_input("Nome do Projeto", value=default_name, key=f"project_name_{form_key}")
    description = st.text_area("Descrição", value=default_description, key=f"project_desc_{form_key}")
    
    # Dropdown de clientes
    client_options = clients_df["name"].tolist() if not clients_df.empty else []
    if client_options:
        if is_edit and default_client in client_options:
            default_client_index = client_options.index(default_client)
        else:
            default_client_index = 0
        client = st.selectbox("Cliente", client_options, index=default_client_index, key=f"project_client_{form_key}")
    else:
        st.error("Nenhum cliente cadastrado. Cadastre um cliente primeiro.")
        return None
    
    # Selectbox para tipo de projeto com validação segura
    if default_type in project_types:
        type_index = project_types.index(default_type)
    else:
        type_index = 0  # Default para "Desenvolvimento"
        
    project_type = st.selectbox("Tipo", 
        project_types,
        index=type_index,
        key=f"project_type_{form_key}")
    
    # Datas
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data de Início", value=default_start_date, key=f"start_date_{form_key}")
    with col2:
        end_date = st.date_input("Data de Término", value=default_end_date, key=f"end_date_{form_key}")
    
    # Interface baseada no modo de cálculo (recebido como parâmetro)
    st.subheader(f"💰 {calculation_mode}")
    
    if calculation_mode == "Rate × Horas = Total":
        col3, col4 = st.columns(2)
        with col3:
            # Selectbox para rates das equipes
            rate_options = ["Personalizado"] + [f"{team} (€{rate:.2f})" for team, rate in team_rates.items()]
            selected_rate_option = st.selectbox(
                "Selecionar Rate da Equipe",
                rate_options,
                help="Escolha um rate pré-definido ou 'Personalizado' para inserir manualmente",
                key=f"rate_option_{form_key}"
            )
            
            if selected_rate_option == "Personalizado":
                hourly_rate = st.number_input(
                    "Valor Hora (€)",
                    min_value=0.0,
                    value=default_hourly_rate,
                    step=1.0,
                    format="%.2f",
                    key=f"hourly_rate_{form_key}"
                )
            else:
                # Extrair o rate do texto selecionado
                team_name = selected_rate_option.split(" (€")[0]
                hourly_rate = team_rates[team_name]
                st.number_input(
                    "Valor Hora (€)",
                    value=hourly_rate,
                    disabled=True,
                    format="%.2f",
                    key=f"hourly_rate_disabled_{form_key}"
                )
        
        with col4:
            total_hours = st.number_input(
                "Quantidade de Horas",
                min_value=0.0,
                value=default_total_hours,
                step=1.0,
                format="%.2f",
                key=f"total_hours_{form_key}"
            )
        
        total_cost = hourly_rate * total_hours if hourly_rate and total_hours else 0.0
        st.success(f"💰 **Custo Total Calculado: €{total_cost:,.2f}**")
        
    elif calculation_mode == "Total ÷ Rate = Horas":
        col3, col4 = st.columns(2)
        with col3:
            total_cost = st.number_input(
                "Custo Total (€)",
                min_value=0.0,
                value=default_total_cost,
                step=1.0,
                format="%.2f",
                key=f"total_cost_{form_key}"
            )
        
        with col4:
            # Selectbox para rates das equipes
            rate_options = ["Personalizado"] + [f"{team} (€{rate:.2f})" for team, rate in team_rates.items()]
            selected_rate_option = st.selectbox(
                "Selecionar Rate da Equipe",
                rate_options,
                help="Escolha um rate pré-definido ou 'Personalizado' para inserir manualmente",
                key=f"rate_option2_{form_key}"
            )
            
            if selected_rate_option == "Personalizado":
                hourly_rate = st.number_input(
                    "Valor Hora (€)",
                    min_value=0.01,
                    value=max(default_hourly_rate, 0.01),
                    step=1.0,
                    format="%.2f",
                    key=f"hourly_rate2_{form_key}"
                )
            else:
                # Extrair o rate do texto selecionado
                team_name = selected_rate_option.split(" (€")[0]
                hourly_rate = team_rates[team_name]
                st.number_input(
                    "Valor Hora (€)",
                    value=hourly_rate,
                    disabled=True,
                    format="%.2f",
                    key=f"hourly_rate2_disabled_{form_key}"
                )
        
        # Calcular horas
        total_hours = total_cost / hourly_rate if hourly_rate > 0 else 0
        st.success(f"⏱️ **Horas Calculadas: {total_hours:,.2f}h**")
        
    elif calculation_mode == "Total ÷ Horas = Rate":
        col3, col4 = st.columns(2)
        with col3:
            total_cost = st.number_input(
                "Custo Total (€)",
                min_value=0.0,
                value=default_total_cost,
                step=1.0,
                format="%.2f",
                key=f"total_cost3_{form_key}"
            )
        
        with col4:
            total_hours = st.number_input(
                "Quantidade de Horas",
                min_value=0.01,
                value=max(default_total_hours, 0.01),
                step=1.0,
                format="%.2f",
                key=f"total_hours3_{form_key}"
            )
        
        # Calcular rate
        hourly_rate = total_cost / total_hours if total_hours > 0 else 0
        st.success(f"💶 **Rate Calculado: €{hourly_rate:,.2f}/hora**")
    
    # Status do projeto
    status = st.selectbox("Status", 
        ["active", "inactive", "completed"],
        index=["active", "inactive", "completed"].index(default_status),
        key=f"project_status_{form_key}")
    
    # Seção de migração de dados
    st.subheader("📊 Dados Migrados de Outro Sistema")
    st.info("💡 Use esta seção para registrar horas e custos já realizados em outros sistemas")
    
    col5, col6 = st.columns(2)
    with col5:
        migrated_hours = st.number_input(
            "Horas Realizadas Migradas",
            min_value=0.0,
            value=default_migrated_hours,
            step=1.0,
            format="%.2f",
            help="Horas já trabalhadas em outro sistema",
            key=f"migrated_hours_{form_key}"
        )
    
    with col6:
        migrated_cost = st.number_input(
            "Custo Realizado Migrado (€)",
            min_value=0.0,
            value=default_migrated_cost,
            step=1.0,
            format="%.2f",
            help="Custo já realizado em outro sistema",
            key=f"migrated_cost_{form_key}"
        )
    
    # Resumo do projeto
    st.subheader("📋 Resumo do Projeto")
    
    col7, col8, col9 = st.columns(3)
    with col7:
        st.metric("Valor Hora", f"€{hourly_rate:,.2f}")
    with col8:
        st.metric("Total de Horas", f"{total_hours:,.2f}h")
    with col9:
        st.metric("Custo Total", f"€{total_cost:,.2f}")
    
    if migrated_hours > 0 or migrated_cost > 0:
        st.subheader("📈 Dados Migrados")
        col10, col11 = st.columns(2)
        with col10:
            st.metric("Horas Migradas", f"{migrated_hours:,.2f}h")
        with col11:
            st.metric("Custo Migrado", f"€{migrated_cost:,.2f}")
    
    return {
        'name': name,
        'description': description,
        'client': client,
        'project_type': project_type,
        'start_date': start_date,
        'end_date': end_date,
        'hourly_rate': hourly_rate,
        'total_hours': total_hours,
        'total_cost': total_cost,
        'status': status,
        'migrated_hours': migrated_hours,
        'migrated_cost': migrated_cost,
        'clients_df': clients_df
    }

def main():
    st.title("Configurações do Sistema")
    
    # Menu para diferentes entidades
    menu = st.sidebar.selectbox(
        "Configurar",
        ["Utilizadores", "Grupos", "Clientes", "Projetos", "Rates", "Categorias de Tarefas"]
    )
    
    # Inicialização dos gerenciadores
    if menu == "Utilizadores":
        current_manager = UserManager()
    elif menu == "Grupos":
        current_manager = GroupManager()
    elif menu == "Clientes":
        current_manager = ClientManager()
    elif menu == "Projetos":
        current_manager = ProjectManager()
    elif menu == "Rates":
        rate_page()
        return
    elif menu == "Categorias de Tarefas":
        current_manager = TaskCategoryManager()
    
    # Interface CRUD
    st.header(menu)
    
    # Botões de ação
    action = st.radio("Ação", ["Listar", "Criar", "Editar", "Excluir"])
    
    if action == "Listar":
        df = current_manager.read()
        
        # Personalizar visualização para Projetos
        if menu == "Projetos" and not df.empty:
            # Carregar dados de clientes para exibir nomes
            clients_manager = ClientManager()
            clients_df = clients_manager.read()
            
            # Adicionar nomes dos clientes
            projects_with_clients = df.merge(
                clients_df[['client_id', 'name']],
                on='client_id',
                how='left'
            )
            
            # Ordenar por status (ativos primeiro) e depois por nome
            projects_with_clients = projects_with_clients.sort_values(
                by=['status', 'project_name'],
                ascending=[False, True]
            )
            
            # Colunas para exibição
            display_columns = [
                'project_id', 'project_name', 'name', 'project_type',
                'start_date', 'end_date', 'hourly_rate', 'total_hours',
                'total_cost', 'status'
            ]
            
            # Adicionar colunas de migração se existirem
            if 'horas_realizadas_mig' in projects_with_clients.columns:
                display_columns.append('horas_realizadas_mig')
            if 'custo_realizado_mig' in projects_with_clients.columns:
                display_columns.append('custo_realizado_mig')
            
            # Filtrar apenas colunas que existem
            existing_columns = [col for col in display_columns if col in projects_with_clients.columns]
            
            # Renomear colunas para exibição
            column_names = {
                'project_id': 'ID',
                'project_name': 'Projeto',
                'name': 'Cliente',
                'project_type': 'Tipo',
                'start_date': 'Data Início',
                'end_date': 'Data Término',
                'hourly_rate': 'Valor Hora (€)',
                'total_hours': 'Quantidade de Horas',
                'total_cost': 'Custo Total (€)',
                'status': 'Status',
                'horas_realizadas_mig': 'Horas Migradas',
                'custo_realizado_mig': 'Custo Migrado (€)'
            }
            
            # Exibir dados formatados
            display_df = projects_with_clients[existing_columns].rename(columns=column_names)
            
            # Função de formatação segura para valores monetários
            def safe_currency_format(x):
                try:
                    if pd.isna(x) or x == '' or x is None:
                        return '€0,00'
                    return f'€{float(x):,.2f}'
                except (ValueError, TypeError):
                    return str(x)
            
            # Função de formatação segura para horas
            def safe_hours_format(x):
                try:
                    if pd.isna(x) or x == '' or x is None:
                        return '0,00'
                    return f'{float(x):,.2f}'
                except (ValueError, TypeError):
                    return str(x)
            
            # Função de formatação segura para datas
            def safe_date_format(x):
                try:
                    if pd.isna(x) or x == '' or x is None:
                        return ''
                    return pd.to_datetime(x).strftime('%d/%m/%Y')
                except (ValueError, TypeError):
                    return str(x)
            
            # Formatação
            format_dict = {
                'Valor Hora (€)': safe_currency_format,
                'Quantidade de Horas': safe_hours_format,
                'Custo Total (€)': safe_currency_format,
                'Data Início': safe_date_format,
                'Data Término': safe_date_format
            }
            
            # Adicionar formatação para campos migrados se existirem
            if 'Horas Migradas' in display_df.columns:
                format_dict['Horas Migradas'] = safe_hours_format
            if 'Custo Migrado (€)' in display_df.columns:
                format_dict['Custo Migrado (€)'] = safe_currency_format
            
            st.dataframe(display_df.style.format(format_dict))
        else:
            st.dataframe(df)
    
    elif action == "Criar":
        # Tratar projetos separadamente
        if menu == "Projetos":
            # TUDO para projetos fica FORA do form
            st.subheader("💰 Modo de Cálculo")
            calculation_modes = [
                "Rate × Horas = Total",
                "Total ÷ Rate = Horas", 
                "Total ÷ Horas = Rate"
            ]
            
            calculation_mode = st.selectbox(
                "Escolha como deseja calcular os valores do projeto",
                calculation_modes,
                key="calc_mode_create"
            )
            
            # Form específico para projetos
            with st.form("create_project_form"):
                project_data = render_project_form(is_edit=False, calculation_mode=calculation_mode, form_key="create")
                
                col1, col2 = st.columns(2)
                with col1:
                    preview_button = st.form_submit_button("🔍 Pré-visualizar", use_container_width=True)
                with col2:
                    create_button = st.form_submit_button("✅ Criar Projeto", use_container_width=True, type="primary")
                
                # Mostrar pré-visualização
                if preview_button:
                    if project_data is None:
                        st.error("Erro ao processar dados do projeto.")
                        st.stop()
                    
                    # Validações básicas para pré-visualização
                    if not project_data['name'] or not project_data['client']:
                        st.error("Nome do projeto e cliente são obrigatórios.")
                        st.stop()
                    
                    if project_data['start_date'] >= project_data['end_date']:
                        st.error("A data de término deve ser posterior à data de início.")
                        st.stop()
                    
                    # Mostrar resumo para confirmação
                    st.success("✅ Pré-visualização do Projeto")
                    st.info(f"""
                    **📋 Resumo do Projeto:**
                    - **Nome**: {project_data['name']}
                    - **Cliente**: {project_data['client']}
                    - **Tipo**: {project_data['project_type']}
                    - **Período**: {project_data['start_date'].strftime('%d/%m/%Y')} até {project_data['end_date'].strftime('%d/%m/%Y')}
                    - **Valor/Hora**: €{project_data['hourly_rate']:,.2f}
                    - **Total de Horas**: {project_data['total_hours']:,.2f}h
                    - **Custo Total**: €{project_data['total_cost']:,.2f}
                    - **Status**: {project_data['status'].title()}
                    
                    {f"**📊 Dados Migrados:**" if project_data['migrated_hours'] > 0 or project_data['migrated_cost'] > 0 else ""}
                    {f"- Horas Migradas: {project_data['migrated_hours']:,.2f}h" if project_data['migrated_hours'] > 0 else ""}
                    {f"- Custo Migrado: €{project_data['migrated_cost']:,.2f}" if project_data['migrated_cost'] > 0 else ""}
                    """)
                    st.warning("⚠️ Revise os dados acima e clique em 'Criar Projeto' para confirmar a criação.")
                
                # Criar projeto
                if create_button:
                    if project_data is None:
                        st.error("Erro ao processar dados do projeto.")
                        st.stop()
                    
                    # Validações
                    if not project_data['name'] or not project_data['client']:
                        st.error("Nome do projeto e cliente são obrigatórios.")
                        st.stop()
                    
                    if project_data['start_date'] >= project_data['end_date']:
                        st.error("A data de término deve ser posterior à data de início.")
                        st.stop()
                    
                    # Obter client_id
                    client_match = project_data['clients_df'][project_data['clients_df']["name"] == project_data['client']]
                    if not client_match.empty:
                        client_id = client_match["client_id"].iloc[0]
                        client_group = client_match["group_id"].iloc[0]
                    else:
                        st.error(f"Cliente '{project_data['client']}' não encontrado")
                        st.stop()
                    
                    data = {
                        "project_name": project_data['name'],
                        "project_description": project_data['description'],
                        "client_id": int(client_id),
                        "group_id": int(client_group),
                        "project_type": project_data['project_type'],
                        "start_date": project_data['start_date'].isoformat(),
                        "end_date": project_data['end_date'].isoformat(),
                        "hourly_rate": float(project_data['hourly_rate']),
                        "total_hours": float(project_data['total_hours']),
                        "total_cost": float(project_data['total_cost']),
                        "status": project_data['status'],
                        "horas_realizadas_mig": float(project_data['migrated_hours']),
                        "custo_realizado_mig": float(project_data['migrated_cost'])
                    }
                    current_manager.create(data)
                    st.success("Projeto criado com sucesso!")
                    st.toast("✅ Item criado!")
        
        else:
            # Form para outros tipos (Usuários, Grupos, Clientes, Categorias)
            with st.form("create_form"):
                # Campos específicos para cada tipo
                if menu == "Utilizadores":
                    first_name = st.text_input("Nome*", help="Campo obrigatório")
                    last_name = st.text_input("Sobrenome*", help="Campo obrigatório")
                    email = st.text_input("Email*", help="Campo obrigatório")
                    password = st.text_input("Senha*", type="password", help="Campo obrigatório")
                    role = st.selectbox("Papel*", ["admin", "leader", "user"], help="Campo obrigatório")
                    
                    # Carregar grupos do banco de dados
                    groups_manager = GroupManager()
                    groups_df = groups_manager.read()
                    groups = st.multiselect("Grupos*", groups_df["group_name"].tolist(), help="Campo obrigatório")
                    
                    active = st.checkbox("Ativo", value=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        preview_user = st.form_submit_button("🔍 Pré-visualizar", use_container_width=True)
                    with col2:
                        create_user = st.form_submit_button("✅ Criar Usuário", use_container_width=True, type="primary")
                    
                    # Pré-visualização do usuário
                    if preview_user:
                        # Validar campos obrigatórios
                        if not first_name or not last_name or not email or not password or not groups:
                            st.error("Todos os campos marcados com * são obrigatórios.")
                            st.stop()
                        
                        # Verificar se o email já existe
                        existing_users = current_manager.read()
                        if not existing_users.empty and email in existing_users['email'].values:
                            st.error(f"O email {email} já está registrado. Por favor, use outro email.")
                            st.stop()
                        
                        # Mostrar resumo
                        st.success("✅ Pré-visualização do Usuário")
                        st.info(f"""
                        **👤 Resumo do Usuário:**
                        - **Nome Completo**: {first_name} {last_name}
                        - **Email**: {email}
                        - **Papel**: {role.title()}
                        - **Grupos**: {', '.join(groups)}
                        - **Status**: {'Ativo' if active else 'Inativo'}
                        """)
                        st.warning("⚠️ Revise os dados acima e clique em 'Criar Usuário' para confirmar a criação.")
                    
                    if create_user:
                        # Validar campos obrigatórios
                        if not first_name or not last_name or not email or not password or not groups:
                            st.error("Todos os campos marcados com * são obrigatórios.")
                            st.stop()
                        
                        # Verificar se o email já existe
                        existing_users = current_manager.read()
                        if not existing_users.empty and email in existing_users['email'].values:
                            st.error(f"O email {email} já está registrado. Por favor, use outro email.")
                            st.stop()
                        
                        # Adicionar timestamp atual
                        current_time = datetime.now().isoformat()
                        
                        data = {
                            "First_Name": first_name,
                            "Last_Name": last_name,
                            "email": email,
                            "password": unified_hash_password(password),
                            "role": role,
                            "groups": str(groups),
                            "active": active,
                            "created_at": current_time,
                            "updated_at": current_time
                        }
                        current_manager.create(data)
                        st.success("Usuário criado com sucesso!")
                        st.toast("✅ Item criado!")
                
                elif menu == "Grupos":
                    name = st.text_input("Nome do Grupo")
                    active = st.checkbox("Ativo", value=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        preview_group = st.form_submit_button("🔍 Pré-visualizar", use_container_width=True)
                    with col2:
                        create_group = st.form_submit_button("✅ Criar Grupo", use_container_width=True, type="primary")
                    
                    # Pré-visualização do grupo
                    if preview_group:
                        if not name:
                            st.error("Nome do grupo é obrigatório.")
                            st.stop()
                        
                        st.success("✅ Pré-visualização do Grupo")
                        st.info(f"""
                        **👥 Resumo do Grupo:**
                        - **Nome**: {name}
                        - **Status**: {'Ativo' if active else 'Inativo'}
                        """)
                        st.warning("⚠️ Revise os dados acima e clique em 'Criar Grupo' para confirmar a criação.")
                    
                    if create_group:
                        if not name:
                            st.error("Nome do grupo é obrigatório.")
                            st.stop()
                        
                        data = {
                            "group_name": name,
                            "active": active
                        }
                        current_manager.create(data)
                        st.success("Grupo criado com sucesso!")
                        st.toast("✅ Item criado!")
                
                elif menu == "Clientes":
                    name = st.text_input("Nome do Cliente")
                    
                    # Carregar grupos do banco de dados
                    groups_manager = GroupManager()
                    groups_df = groups_manager.read()
                    
                    if groups_df.empty:
                        st.error("Nenhum grupo cadastrado. Cadastre um grupo primeiro.")
                        st.stop()
                    
                    group = st.selectbox("Grupo", groups_df["group_name"].tolist())
                    contact = st.text_input("Contato")
                    email = st.text_input("Email")
                    active = st.checkbox("Ativo", value=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        preview_client = st.form_submit_button("🔍 Pré-visualizar", use_container_width=True)
                    with col2:
                        create_client = st.form_submit_button("✅ Criar Cliente", use_container_width=True, type="primary")
                    
                    # Pré-visualização do cliente
                    if preview_client:
                        if not name:
                            st.error("Nome do cliente é obrigatório.")
                            st.stop()
                        
                        st.success("✅ Pré-visualização do Cliente")
                        st.info(f"""
                        **🏢 Resumo do Cliente:**
                        - **Nome**: {name}
                        - **Grupo**: {group}
                        - **Contato**: {contact if contact else 'Não informado'}
                        - **Email**: {email if email else 'Não informado'}
                        - **Status**: {'Ativo' if active else 'Inativo'}
                        """)
                        st.warning("⚠️ Revise os dados acima e clique em 'Criar Cliente' para confirmar a criação.")
                    
                    if create_client:
                        if not name:
                            st.error("Nome do cliente é obrigatório.")
                            st.stop()
                        
                        # Encontrar o group_id correspondente
                        group_match = groups_df[groups_df["group_name"] == group]
                        if not group_match.empty:
                            group_id = group_match["id"].iloc[0]
                        else:
                            st.error(f"Grupo '{group}' não encontrado")
                            st.stop()
                        
                        data = {
                            "name": name,
                            "group_id": int(group_id),  # Garantir que seja inteiro
                            "contact": contact,
                            "email": email,
                            "active": active
                        }
                        current_manager.create(data)
                        st.success("Cliente criado com sucesso!")
                        st.toast("✅ Item criado!")
                
                elif menu == "Categorias de Tarefas":
                    category = st.text_input("Nome da Categoria")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        preview_category = st.form_submit_button("🔍 Pré-visualizar", use_container_width=True)
                    with col2:
                        create_category = st.form_submit_button("✅ Criar Categoria", use_container_width=True, type="primary")
                    
                    # Pré-visualização da categoria
                    if preview_category:
                        if not category:
                            st.error("Nome da categoria é obrigatório.")
                            st.stop()
                        
                        st.success("✅ Pré-visualização da Categoria")
                        st.info(f"""
                        **📋 Resumo da Categoria:**
                        - **Nome**: {category}
                        """)
                        st.warning("⚠️ Revise os dados acima e clique em 'Criar Categoria' para confirmar a criação.")
                    
                    if create_category:
                        if not category:
                            st.error("Nome da categoria é obrigatório.")
                            st.stop()
                        
                        data = {
                            "task_category": category
                        }
                        current_manager.create(data)
                        st.success("Categoria criada com sucesso!")
                        st.toast("✅ Item criado!")
    
    elif action == "Editar":
        items = current_manager.read()
        
        if items.empty:
            st.warning(f"Não há itens para editar.")
            return
        
        # Seleção do item para editar com base no tipo
        id_column = {
            "Utilizadores": "user_id",
            "Grupos": "id",
            "Clientes": "client_id",
            "Projetos": "project_id",
            "Categorias de Tarefas": "task_category_id"
        }.get(menu)
        
        name_column = {
            "Utilizadores": lambda x: f"{x['First_Name']} {x['Last_Name']}",
            "Grupos": "group_name",
            "Clientes": "name",
            "Projetos": "project_name",
            "Categorias de Tarefas": "task_category"
        }.get(menu)

        # Criar lista de opções
        options = []
        for _, item in items.iterrows():
            if id_column in item and id_column:
                item_id = item[id_column]
                if callable(name_column):
                    display_name = name_column(item)
                elif name_column in item:
                    display_name = item[name_column]
                else:
                    display_name = f"ID: {item_id}"
                options.append({"id": item_id, "name": display_name})
        
        if not options:
            st.warning(f"Não há opções disponíveis para seleção.")
            return
            
        selected_option = st.selectbox(
            f"Selecionar {menu[:-1].lower()}", 
            options,
            format_func=lambda x: x["name"]
        )
        
        if selected_option:
            if isinstance(selected_option, dict):
                item_id = selected_option["id"]
            else:
                item_id = selected_option
            item = items[items[id_column] == item_id]
            
            if not item.empty:
                item = item.iloc[0]
                
                # Tratar projetos separadamente
                if menu == "Projetos":
                    # TUDO para projetos fica FORA do form
                    st.subheader("💰 Modo de Cálculo")
                    calculation_modes = [
                        "Rate × Horas = Total",
                        "Total ÷ Rate = Horas", 
                        "Total ÷ Horas = Rate"
                    ]
                    
                    calculation_mode = st.selectbox(
                        "Escolha como deseja calcular os valores do projeto",
                        calculation_modes,
                        key="calc_mode_edit"
                    )
                    
                    # Form específico para projetos
                    with st.form("edit_project_form"):
                        project_data = render_project_form(is_edit=True, item=item, calculation_mode=calculation_mode, form_key="edit")
                        
                        if st.form_submit_button("Atualizar"):
                            if project_data is None:
                                st.error("Erro ao processar dados do projeto.")
                                st.stop()
                            
                            # Validações
                            if not project_data['name'] or not project_data['client']:
                                st.error("Nome do projeto e cliente são obrigatórios.")
                                st.stop()
                            
                            if project_data['start_date'] >= project_data['end_date']:
                                st.error("A data de término deve ser posterior à data de início.")
                                st.stop()
                            
                            # Obter client_id
                            client_match = project_data['clients_df'][project_data['clients_df']["name"] == project_data['client']]
                            if not client_match.empty:
                                client_id = client_match["client_id"].iloc[0]
                                client_group = client_match["group_id"].iloc[0]
                            else:
                                st.error(f"Cliente '{project_data['client']}' não encontrado")
                                st.stop()

                            
                                
                            data = {
                                "project_name": project_data['name'],
                                "project_description": project_data['description'],
                                "client_id": int(client_id),
                                "group_id": int(client_group),
                                "project_type": project_data['project_type'],
                                "start_date": project_data['start_date'].isoformat(),
                                "end_date": project_data['end_date'].isoformat(),
                                "hourly_rate": float(project_data['hourly_rate']),
                                "total_hours": float(project_data['total_hours']),
                                "total_cost": float(project_data['total_cost']),
                                "status": project_data['status'],
                                "horas_realizadas_mig": float(project_data['migrated_hours']),
                                "custo_realizado_mig": float(project_data['migrated_cost'])
                            }
                            current_manager.update(item_id, data)
                            st.success("Projeto atualizado com sucesso!")
                
                else:
                    # Form para outros tipos (Usuários, Grupos, Clientes, Categorias)
                    with st.form("edit_form"):
                        if menu == "Utilizadores":
                            first_name = st.text_input("Nome*", item["First_Name"], help="Campo obrigatório")
                            last_name = st.text_input("Sobrenome*", item["Last_Name"], help="Campo obrigatório")
                            email = st.text_input("Email*", item["email"], help="Campo obrigatório")
                            password = st.text_input("Senha", type="password", help="Deixe em branco para manter a senha atual")
                            
                            # Corrigido o tratamento do role
                            current_role = str(item["role"]).lower()
                            if current_role not in ["admin", "leader", "user"]:
                                current_role = "user"
                                
                            role = st.selectbox("Papel*", 
                                ["admin", "leader", "user"],
                                index=["admin", "leader", "user"].index(current_role),
                                help="Campo obrigatório")
                            
                            # Carregar grupos do banco de dados
                            groups_manager = GroupManager()
                            groups_df = groups_manager.read()
                            
                            # Processar os grupos atuais
                            try:
                                current_groups = eval(item["groups"]) if isinstance(item["groups"], str) else []
                            except:
                                current_groups = []
                            
                            groups = st.multiselect("Grupos*", 
                                groups_df["group_name"].tolist(),
                                default=current_groups,
                                help="Campo obrigatório")
                            
                            active = st.checkbox("Ativo", bool(item["active"]))
                            
                            if st.form_submit_button("Atualizar"):
                                # Validar campos obrigatórios
                                if not first_name or not last_name or not email or not groups:
                                    st.error("Todos os campos marcados com * são obrigatórios.")
                                    st.stop()
                                    
                                # Verificar se o email já existe para outros usuários
                                existing_users = current_manager.read()
                                if not existing_users.empty:
                                    duplicate_email = existing_users[(existing_users['email'] == email) & 
                                                                (existing_users[id_column] != item_id)]
                                    if not duplicate_email.empty:
                                        st.error(f"O email {email} já está sendo usado por outro usuário.")
                                        st.stop()
                                
                                # Adicionar timestamp atual
                                current_time = datetime.now().isoformat()
                                
                                # Criar dicionário de dados
                                data = {
                                    "First_Name": first_name,
                                    "Last_Name": last_name,
                                    "email": email,
                                    "role": role,
                                    "groups": str(groups),
                                    "active": 1 if active else 0,
                                    "updated_at": current_time
                                }
                                
                                # Adicionar senha apenas se foi fornecida
                                if password and password.strip():
                                    data["password"] = unified_hash_password(password)
                                    
                                current_manager.update(item_id, data)
                                st.success("Usuário atualizado com sucesso!")
                        
                        elif menu == "Grupos":
                            name = st.text_input("Nome do Grupo", item["group_name"])
                            active = st.checkbox("Ativo", bool(item["active"]))
                            
                            if st.form_submit_button("Atualizar"):
                                data = {
                                    "group_name": name,
                                    "active": 1 if active else 0
                                }
                                current_manager.update(item_id, data)
                                st.success("Grupo atualizado com sucesso!")
                        
                        elif menu == "Clientes":
                            name = st.text_input("Nome do Cliente", item["name"])
                            
                            # Carregar grupos do banco de dados
                            groups_manager = GroupManager()
                            groups_df = groups_manager.read()
                            
                            # Encontrar o grupo atual
                            group_matches = groups_df[groups_df["id"] == item["group_id"]]
                            current_group = group_matches["group_name"].iloc[0] if not group_matches.empty else ""
                            
                            # Se não encontrarmos o grupo, usamos o primeiro da lista ou string vazia
                            default_index = 0
                            group_options = groups_df["group_name"].tolist()
                            
                            if current_group in group_options:
                                default_index = group_options.index(current_group)
                            
                            group = st.selectbox("Grupo", group_options, 
                                index=default_index)
                            
                            contact = st.text_input("Contato", item["contact"] if not pd.isna(item["contact"]) else "")
                            email = st.text_input("Email", item["email"] if not pd.isna(item["email"]) else "")
                            active = st.checkbox("Ativo", bool(item["active"]))
                            
                            if st.form_submit_button("Atualizar"):
                                # Obter o ID do grupo selecionado
                                group_match = groups_df[groups_df["group_name"] == group]
                                if not group_match.empty:
                                    group_id = group_match["id"].iloc[0]
                                else:
                                    st.error(f"Grupo '{group}' não encontrado")
                                    st.stop()
                                    
                                data = {
                                    "name": name,
                                    "group_id": int(group_id),  # Garantir que seja inteiro
                                    "contact": contact,
                                    "email": email,
                                    "active": 1 if active else 0
                                }
                                current_manager.update(item_id, data)
                                st.success("Cliente atualizado com sucesso!")
                        
                        elif menu == "Categorias de Tarefas":
                            category = st.text_input("Nome da Categoria", item["task_category"])
                            
                            if st.form_submit_button("Atualizar"):
                                data = {
                                    "task_category": category
                                }
                                current_manager.update(item_id, data)
                                st.success("Categoria atualizada com sucesso!")
    
    elif action == "Excluir":
        items = current_manager.read()
        
        if items.empty:
            st.warning(f"Não há itens para excluir.")
            return
        
        # Determinar a coluna de ID e nome com base no menu
        id_column = {
            "Utilizadores": "user_id",
            "Grupos": "id",
            "Clientes": "client_id",
            "Projetos": "project_id",
            "Categorias de Tarefas": "task_category_id"
        }.get(menu)
        
        name_column = {
            "Utilizadores": lambda x: f"{x['First_Name']} {x['Last_Name']}",
            "Grupos": "group_name",
            "Clientes": "name",
            "Projetos": "project_name",
            "Categorias de Tarefas": "task_category"
        }.get(menu)

        # Criar lista de opções
        options = []
        for _, item in items.iterrows():
            if id_column in item and id_column:
                item_id = item[id_column]
                if callable(name_column):
                    display_name = name_column(item)
                elif name_column in item:
                    display_name = item[name_column]
                else:
                    display_name = f"ID: {item_id}"
                options.append({"id": item_id, "name": display_name})
        
        if not options:
            st.warning(f"Não há opções disponíveis para exclusão.")
            return
            
        selected_option = st.selectbox(
            f"Selecionar {menu[:-1].lower()} para excluir", 
            options,
            format_func=lambda x: x["name"]
        )
        
        if selected_option:
            item_id = selected_option["id"]
            
            if st.button(f"Confirmar Exclusão de {selected_option['name']}"):
                current_manager.delete(item_id)
                st.success(f"Item excluído com sucesso!")

if __name__ == "__main__":
    main()