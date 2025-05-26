# billing_manager.py

import streamlit as st
import pandas as pd
from datetime import datetime
from database_manager import DatabaseManager

class BillingManager:
    def __init__(self):
        self.db = DatabaseManager()
        
        # Criar tabela se não existir
        self._create_tables()
    
    def _create_tables(self):
        """Cria a tabela de faturas se não existir ou recria se estiver com estrutura incompatível"""
        # Primeiro verificamos se a tabela existe
        check_query = "SELECT name FROM sqlite_master WHERE type='table' AND name='invoices'"
        tables = self.db.query_to_df(check_query)
        
        # Se a tabela existir, verificamos sua estrutura
        if not tables.empty and 'invoices' in tables['name'].values:
            try:
                # Tente fazer uma consulta simples para verificar se a estrutura está correta
                test_query = "SELECT client_id FROM invoices LIMIT 1"
                self.db.query_to_df(test_query)
            except Exception:
                # Se falhar, precisamos recriar a tabela
                print("Tabela invoices existe mas tem estrutura incompatível. Recriando...")
                drop_query = "DROP TABLE invoices"
                self.db.execute_query(drop_query)
                # A tabela será criada abaixo
        
        # Criar a tabela
        create_query = """
        CREATE TABLE IF NOT EXISTS invoices (
            invoice_id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER NOT NULL,
            project_id INTEGER NOT NULL,
            invoice_number TEXT NOT NULL,
            amount REAL NOT NULL,
            issue_date TIMESTAMP NOT NULL,
            payment_date TIMESTAMP NOT NULL,
            payment_method TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(client_id),
            FOREIGN KEY (project_id) REFERENCES projects(project_id)
        );
        """
        self.db.execute_query(create_query)

    def create_invoice(self, data):
        """Cria um novo registro de fatura"""
        # Adicionar timestamps se não existirem
        if 'created_at' not in data:
            data['created_at'] = datetime.now().isoformat()
        if 'updated_at' not in data:
            data['updated_at'] = datetime.now().isoformat()
        
        # Inserir no banco de dados
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        
        query = f"INSERT INTO invoices ({columns}) VALUES ({placeholders})"
        cursor = self.db.execute_query(query, tuple(data.values()))
        
        return cursor.lastrowid
    
    def get_invoice(self, invoice_id=None):
        """Obtém faturas com informações do cliente e projeto"""
        try:
            # Verificar se a tabela existe primeiro
            table_check = "SELECT name FROM sqlite_master WHERE type='table' AND name='invoices'"
            tables = self.db.query_to_df(table_check)
            
            if tables.empty or 'invoices' not in tables['name'].values:
                # Tabela não existe, criar e retornar DataFrame vazio
                self._create_tables()
                return pd.DataFrame()
                
            if invoice_id:
                query = """
                SELECT i.*, c.name as client_name, p.project_name
                FROM invoices i
                JOIN clients c ON i.client_id = c.client_id
                JOIN projects p ON i.project_id = p.project_id
                WHERE i.invoice_id = ?
                """
                return self.db.query_to_df(query, (invoice_id,))
            else:
                query = """
                SELECT i.*, c.name as client_name, p.project_name
                FROM invoices i
                JOIN clients c ON i.client_id = c.client_id
                JOIN projects p ON i.project_id = p.project_id
                ORDER BY i.payment_date DESC
                """
                return self.db.query_to_df(query)
        except Exception:
            # Em caso de erro, retorna DataFrame vazio
            return pd.DataFrame()
    
    def update_invoice(self, invoice_id, data):
        """Atualiza um registro de fatura"""
        # Adicionar timestamp atualizado
        data['updated_at'] = datetime.now().isoformat()
        
        # Preparar a query
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        
        query = f"UPDATE invoices SET {set_clause} WHERE invoice_id = ?"
        
        params = list(data.values())
        params.append(invoice_id)
        
        self.db.execute_query(query, params)
        return True
    
    def delete_invoice(self, invoice_id):
        """Exclui um registro de fatura"""
        query = "DELETE FROM invoices WHERE invoice_id = ?"
        self.db.execute_query(query, (invoice_id,))
        return True
    
    def get_client_total(self, client_id):
        """Obtém o total faturado para um cliente"""
        # Primeiro verifica se a tabela existe
        try:
            query = "SELECT SUM(amount) as total FROM invoices WHERE client_id = ?"
            result = self.db.query_to_df(query, (client_id,))
            if result.empty or pd.isna(result['total'].iloc[0]):
                return 0
            return float(result['total'].iloc[0])
        except Exception:
            # Se houver erro (tabela não existe), retorna 0
            return 0
    
    def get_project_total(self, project_id):
        """Obtém o total faturado para um projeto"""
        # Primeiro verifica se a tabela existe
        try:
            query = "SELECT SUM(amount) as total FROM invoices WHERE project_id = ?"
            result = self.db.query_to_df(query, (project_id,))
            if result.empty or pd.isna(result['total'].iloc[0]):
                return 0
            return float(result['total'].iloc[0])
        except Exception:
            # Se houver erro (tabela não existe), retorna 0
            return 0
    
    def get_invoices_by_period(self, start_date, end_date):
        """Obtém faturas dentro de um período"""
        try:
            # Verificar se a tabela existe primeiro
            table_check = "SELECT name FROM sqlite_master WHERE type='table' AND name='invoices'"
            tables = self.db.query_to_df(table_check)
            
            if tables.empty or 'invoices' not in tables['name'].values:
                # Tabela não existe, criar e retornar DataFrame vazio
                self._create_tables()
                return pd.DataFrame()
                
            query = """
            SELECT i.*, c.name as client_name, p.project_name
            FROM invoices i
            JOIN clients c ON i.client_id = c.client_id
            JOIN projects p ON i.project_id = p.project_id
            WHERE i.payment_date >= ? AND i.payment_date <= ?
            ORDER BY i.payment_date DESC
            """
            return self.db.query_to_df(query, (start_date.isoformat(), end_date.isoformat()))
        except Exception:
            # Em caso de erro, retorna DataFrame vazio
            return pd.DataFrame()

def billing_page():
    """Página de gestão de faturas"""
    st.title("Gestão de Faturas")
    
    # Inicializar gerenciadores
    billing_manager = BillingManager()
    db_manager = DatabaseManager()
    
    # Carregar dados necessários
    clients_df = db_manager.query_to_df("SELECT * FROM clients WHERE active = 1")
    
    # Criar tabs para melhor organização
    tab1, tab2 = st.tabs(["Registar Faturas", "Consultar Faturas"])
    
    with tab1:
        st.header("Registar Nova Fatura")
        
        if clients_df.empty:
            st.warning("Não há clientes ativos cadastrados.")
        else:
            # Seleção de cliente
            client = st.selectbox(
                "Cliente",
                options=clients_df['name'].tolist(),
                key="invoice_client"
            )
            
            if client:
                # Obter ID do cliente
                client_id = int(clients_df[clients_df['name'] == client]['client_id'].iloc[0])
                
                # Carregar projetos do cliente
                projects_df = db_manager.query_to_df(
                    "SELECT * FROM projects WHERE client_id = ? AND status IN ('active', 'Active')",
                    (client_id,)
                )
                
                if projects_df.empty:
                    st.warning("Não há projetos ativos para este cliente.")
                else:
                    # Formulário para registrar nova fatura
                    with st.form(key="new_invoice_form"):
                        st.subheader("Detalhes da Fatura")
                        
                        # Seleção de projeto
                        project = st.selectbox(
                            "Projeto",
                            options=projects_df['project_name'].tolist()
                        )
                        
                        # Obter project_id
                        project_id = int(projects_df[
                            projects_df['project_name'] == project
                        ]['project_id'].iloc[0])
                        
                        # Mostrar faturação total do projeto até agora
                        project_total = billing_manager.get_project_total(project_id)
                        if project_total > 0:
                            st.info(f"Total já faturado neste projeto: €{project_total:,.2f}")
                        
                        # Campo para número da fatura
                        col1, col2 = st.columns(2)
                        with col1:
                            invoice_number = st.text_input(
                                "Número da Fatura",
                                placeholder="Ex: FAT-2025-001"
                            )
                        
                        with col2:
                            amount = st.number_input(
                                "Valor da Fatura (€)",
                                min_value=0.01,
                                step=100.0,
                                format="%.2f"
                            )
                        
                        # Datas
                        col3, col4 = st.columns(2)
                        with col3:
                            issue_date = st.date_input(
                                "Data de Emissão",
                                value=datetime.now().date()
                            )
                        
                        with col4:
                            payment_date = st.date_input(
                                "Data de Pagamento",
                                value=datetime.now().date()
                            )
                        
                        # Método de pagamento
                        payment_method = st.selectbox(
                            "Método de Pagamento",
                            options=["Transferência Bancária", "Cheque", "Cartão", "Outro"]
                        )
                        
                        # Notas adicionais
                        notes = st.text_area(
                            "Observações",
                            placeholder="Informações adicionais sobre a fatura..."
                        )
                        
                        # Botão para submeter
                        submitted = st.form_submit_button("Registar Fatura", type="primary")
                        
                        if submitted:
                            if not invoice_number:
                                st.error("Por favor, preencha o número da fatura.")
                            elif amount <= 0:
                                st.error("O valor da fatura deve ser maior que zero.")
                            else:
                                try:
                                    # Preparar dados
                                    invoice_data = {
                                        'client_id': client_id,
                                        'project_id': project_id,
                                        'invoice_number': invoice_number,
                                        'amount': amount,
                                        'issue_date': issue_date.isoformat(),
                                        'payment_date': payment_date.isoformat(),
                                        'payment_method': payment_method,
                                        'notes': notes
                                    }
                                    
                                    # Criar fatura
                                    invoice_id = billing_manager.create_invoice(invoice_data)
                                    
                                    if invoice_id:
                                        st.success(f"Fatura {invoice_number} registrada com sucesso!")
                                        st.balloons()
                                    else:
                                        st.error("Erro ao registrar fatura.")
                                except Exception as e:
                                    st.error(f"Erro: {str(e)}")
    
    with tab2:
        st.header("Consultar Faturas")
        
        # Opções de filtro
        col1, col2 = st.columns(2)
        
        with col1:
            filter_type = st.selectbox(
                "Filtrar por",
                options=["Todas as Faturas", "Cliente", "Projeto", "Período"]
            )
        
        # Aplicar filtro selecionado
        if filter_type == "Cliente":
            with col2:
                if clients_df.empty:
                    st.warning("Não há clientes cadastrados.")
                    invoices_df = pd.DataFrame()
                else:
                    filter_client = st.selectbox(
                        "Selecionar Cliente",
                        options=clients_df['name'].tolist()
                    )
                    
                    # Obter ID do cliente
                    filter_client_id = int(clients_df[
                        clients_df['name'] == filter_client
                    ]['client_id'].iloc[0])
                    
                                                # Buscar faturas do cliente
                    try:
                        invoices_df = db_manager.query_to_df(
                            """
                            SELECT i.*, c.name as client_name, p.project_name
                            FROM invoices i
                            JOIN clients c ON i.client_id = c.client_id
                            JOIN projects p ON i.project_id = p.project_id
                            WHERE i.client_id = ?
                            ORDER BY i.payment_date DESC
                            """,
                            (filter_client_id,)
                        )
                    except Exception:
                        st.warning("Nenhuma fatura encontrada para este cliente.")
                        invoices_df = pd.DataFrame()
        
        elif filter_type == "Projeto":
            with col2:
                # Primeiro seleciona o cliente
                if clients_df.empty:
                    st.warning("Não há clientes cadastrados.")
                    invoices_df = pd.DataFrame()
                else:
                    filter_client = st.selectbox(
                        "Selecionar Cliente",
                        options=clients_df['name'].tolist()
                    )
                    
                    # Obter ID do cliente
                    filter_client_id = int(clients_df[
                        clients_df['name'] == filter_client
                    ]['client_id'].iloc[0])
                    
                    # Carregar projetos do cliente
                    filter_projects_df = db_manager.query_to_df(
                        "SELECT * FROM projects WHERE client_id = ?",
                        (filter_client_id,)
                    )
                    
                    if filter_projects_df.empty:
                        st.warning("Não há projetos para este cliente.")
                        invoices_df = pd.DataFrame()
                    else:
                        filter_project = st.selectbox(
                            "Selecionar Projeto",
                            options=filter_projects_df['project_name'].tolist()
                        )
                        
                        # Obter project_id
                        filter_project_id = int(filter_projects_df[
                            filter_projects_df['project_name'] == filter_project
                        ]['project_id'].iloc[0])
                        
                        # Buscar faturas do projeto
                        try:
                            invoices_df = db_manager.query_to_df(
                                """
                                SELECT i.*, c.name as client_name, p.project_name
                                FROM invoices i
                                JOIN clients c ON i.client_id = c.client_id
                                JOIN projects p ON i.project_id = p.project_id
                                WHERE i.project_id = ?
                                ORDER BY i.payment_date DESC
                                """,
                                (filter_project_id,)
                            )
                        except Exception:
                            st.warning("Nenhuma fatura encontrada para este projeto.")
                            invoices_df = pd.DataFrame()
        
        elif filter_type == "Período":
            col2a, col2b = st.columns(2)
            with col2a:
                start_date = st.date_input(
                    "Data Inicial",
                    value=datetime(datetime.now().year, 1, 1).date()
                )
            with col2b:
                end_date = st.date_input(
                    "Data Final",
                    value=datetime.now().date()
                )
            
            # Buscar faturas do período
            invoices_df = billing_manager.get_invoices_by_period(
                start_date,
                end_date
            )
        
        else:  # Todas as Faturas
            # Buscar todas as faturas
            invoices_df = billing_manager.get_invoice()
        
        # Exibir resultados
        if 'invoices_df' in locals() and not invoices_df.empty:
            # Preparar para exibição
            display_df = invoices_df.copy()
            
            # Formatar datas
            display_df['issue_date'] = pd.to_datetime(display_df['issue_date']).dt.strftime('%d/%m/%Y')
            display_df['payment_date'] = pd.to_datetime(display_df['payment_date']).dt.strftime('%d/%m/%Y')
            
            # Selecionar e renomear colunas
            display_df = display_df[[
                'invoice_id', 'invoice_number', 'client_name', 'project_name',
                'amount', 'issue_date', 'payment_date', 'payment_method'
            ]].rename(columns={
                'invoice_id': 'ID',
                'invoice_number': 'Número da Fatura',
                'client_name': 'Cliente',
                'project_name': 'Projeto',
                'amount': 'Valor (€)',
                'issue_date': 'Data de Emissão',
                'payment_date': 'Data de Pagamento',
                'payment_method': 'Método de Pagamento'
            })
            
            # Formatar valores
            display_df['Valor (€)'] = display_df['Valor (€)'].apply(
                lambda x: f"€{float(x):,.2f}"
            )
            
            # Exibir tabela
            st.dataframe(display_df)
            
            # Mostrar total
            total_amount = invoices_df['amount'].sum()
            st.info(f"Total das faturas: **€{total_amount:,.2f}**")
            
            # Opções para editar ou excluir fatura
            st.subheader("Gerenciar Fatura")
            
            # Seleção de fatura
            invoice_options = [
                f"{row['invoice_number']} - {row['client_name']} - {row['project_name']} ({pd.to_datetime(row['payment_date']).strftime('%d/%m/%Y')})"
                for _, row in invoices_df.iterrows()
            ]
            
            invoice_ids = invoices_df['invoice_id'].tolist()
            
            if invoice_options:
                selected_invoice_index = st.selectbox(
                    "Selecionar Fatura",
                    range(len(invoice_options)),
                    format_func=lambda i: invoice_options[i]
                )
                
                selected_invoice_id = invoice_ids[selected_invoice_index]
                selected_invoice = invoices_df[
                    invoices_df['invoice_id'] == selected_invoice_id
                ].iloc[0]
                
                action = st.radio(
                    "Ação",
                    options=["Editar", "Excluir"]
                )
                
                if action == "Editar":
                    with st.form(key="edit_invoice_form"):
                        st.subheader("Editar Fatura")
                        
                        # Campos pré-preenchidos
                        invoice_number = st.text_input(
                            "Número da Fatura",
                            value=selected_invoice['invoice_number']
                        )
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            amount = st.number_input(
                                "Valor da Fatura (€)",
                                min_value=0.01,
                                value=float(selected_invoice['amount']),
                                step=100.0,
                                format="%.2f"
                            )
                        
                        with col2:
                            payment_method = st.selectbox(
                                "Método de Pagamento",
                                options=["Transferência Bancária", "Cheque", "Cartão", "Outro"],
                                index=["Transferência Bancária", "Cheque", "Cartão", "Outro"].index(
                                    selected_invoice['payment_method']
                                ) if selected_invoice['payment_method'] in [
                                    "Transferência Bancária", "Cheque", "Cartão", "Outro"
                                ] else 0
                            )
                        
                        col3, col4 = st.columns(2)
                        with col3:
                            issue_date = st.date_input(
                                "Data de Emissão",
                                value=pd.to_datetime(selected_invoice['issue_date']).date()
                            )
                        
                        with col4:
                            payment_date = st.date_input(
                                "Data de Pagamento",
                                value=pd.to_datetime(selected_invoice['payment_date']).date()
                            )
                        
                        notes = st.text_area(
                            "Observações",
                            value=selected_invoice['notes'] if 'notes' in selected_invoice and selected_invoice['notes'] else ""
                        )
                        
                        # Botão para submeter
                        submitted = st.form_submit_button("Atualizar Fatura", type="primary")
                        
                        if submitted:
                            if not invoice_number:
                                st.error("Por favor, preencha o número da fatura.")
                            elif amount <= 0:
                                st.error("O valor da fatura deve ser maior que zero.")
                            else:
                                try:
                                    # Preparar dados
                                    invoice_data = {
                                        'invoice_number': invoice_number,
                                        'amount': amount,
                                        'issue_date': issue_date.isoformat(),
                                        'payment_date': payment_date.isoformat(),
                                        'payment_method': payment_method,
                                        'notes': notes
                                    }
                                    
                                    # Atualizar fatura
                                    success = billing_manager.update_invoice(
                                        selected_invoice_id,
                                        invoice_data
                                    )
                                    
                                    if success:
                                        st.success(f"Fatura {invoice_number} atualizada com sucesso!")
                                        st.rerun()
                                    else:
                                        st.error("Erro ao atualizar fatura.")
                                except Exception as e:
                                    st.error(f"Erro: {str(e)}")
                
                elif action == "Excluir":
                    st.warning(
                        f"Tem certeza que deseja excluir a fatura {selected_invoice['invoice_number']}?\n"
                        f"Esta ação não pode ser desfeita."
                    )
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("Sim, Excluir", type="primary"):
                            try:
                                # Excluir fatura
                                success = billing_manager.delete_invoice(selected_invoice_id)
                                
                                if success:
                                    st.success(f"Fatura {selected_invoice['invoice_number']} excluída com sucesso!")
                                    st.rerun()
                                else:
                                    st.error("Erro ao excluir fatura.")
                            except Exception as e:
                                st.error(f"Erro: {str(e)}")
                    
                    with col2:
                        if st.button("Cancelar"):
                            st.rerun()
            
            # Exportar dados
            if st.button("Exportar para Excel"):
                # Preparar para exportação
                export_df = invoices_df.copy()
                
                # Formatar datas
                export_df['issue_date'] = pd.to_datetime(export_df['issue_date']).dt.strftime('%d/%m/%Y')
                export_df['payment_date'] = pd.to_datetime(export_df['payment_date']).dt.strftime('%d/%m/%Y')
                
                # Salvar temporariamente
                export_path = 'faturas_export.xlsx'
                export_df.to_excel(export_path, index=False)
                
                # Oferecer para download
                with open(export_path, "rb") as f:
                    bytes_data = f.read()
                    st.download_button(
                        label="Baixar Excel",
                        data=bytes_data,
                        file_name="faturas_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("Nenhuma fatura encontrada com os filtros atuais.")

if __name__ == "__main__":
    billing_page()