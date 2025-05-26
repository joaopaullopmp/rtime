# main.py
import streamlit as st
st.set_page_config(page_title="Sistema de Gestão de Horas", page_icon="⏱️", layout="wide")

import pandas as pd
from pathlib import Path
import hashlib
from rate_manager import RateManager, rate_page
import time
import os
from timesheet import timesheet_page
from project_reports import reports_page
from productivity_reports import team_productivity_page, user_productivity_page
from absence_reports import ausencias_report_page
from unrecorded_hours_reports import horas_nao_registradas_report
from risk_reports import relatorio_projetos_por_risco
from calendar_view import calendar_page
from productivity_dashboard import productivity_dashboard
from app import main as app_main
from auth import Auth
from datetime import datetime
from worked_hours_report import worked_hours_report
# Importe as funções de acesso ao banco de dados
from db_utils import read_table_from_db, save_table_to_db
from database_manager import DatabaseManager, UserManager, ClientManager, ProjectManager, GroupManager
from billing_manager import billing_page
from dashboard import dashboard_page
from dashboard_debug import dashboard_debug
from user_targets_dashboard import user_targets_dashboard
from project_status_email import project_status_email
#from project_phases import project_phases_page
#from phase_progress import phase_progress_dashboard
from executive_dashboard_email import executive_dashboard_email
from collaborator_email_report import collaborator_email_report
from project_email_report import project_email_report
from revenue_email_report import revenue_email_report
#from commercial_meetings_report import commercial_meetings_report
#from auto_project_status_email import project_status_email
from comercial_indicators_email import commercial_indicators_email


def main():
    auth = Auth()
    auth.initialize_session()
    
    if not st.session_state.logged_in:
        # Esconder menu do Streamlit na página de login
        st.markdown("""
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            .block-container {padding-top: 2rem;}
            </style>
        """, unsafe_allow_html=True)
        
        # Centralizar logo e formulário
        col1, col2, col3 = st.columns([1,2,1])
        
        with col2:
            # Logo centralizado
            st.markdown("<br>" * 2, unsafe_allow_html=True)
            st.image("logo.png", width=300)
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Formulário de login
            with st.form("login_form"):
                email = st.text_input("Email")
                password = st.text_input("Senha", type="password")
                st.markdown("<br>", unsafe_allow_html=True)
                submit = st.form_submit_button("Entrar", use_container_width=True)
                
                if submit:
                    success, result = auth.login(email, password)
                    if success:
                        st.session_state.logged_in = True
                        st.session_state.user_info = result
                        st.success("Login realizado com sucesso!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(result)
    
    else:
        # Logo no sidebar com tamanho reduzido
        st.sidebar.image("logo.png", width=250)
        st.sidebar.markdown('<div style="margin-bottom: -10px; margin-top: -15px;"><hr /></div>', unsafe_allow_html=True)
        
        # Informações do Usuário - reduzido espaçamento
        user_info = st.session_state.user_info
        user_groups = eval(user_info['groups']) if isinstance(user_info['groups'], str) else user_info['groups']
        groups_str = user_groups[0] if isinstance(user_groups, list) and user_groups else "Sem grupo"
        
        st.sidebar.markdown(
            f"<div style='margin-bottom: -10px;'>👤 User: <b>{user_info['First_Name']} {user_info['Last_Name']}</b> ({user_info['role'].title()}) <br/>"
            f"👥 Equipa: {groups_str}</div>",
            unsafe_allow_html=True
        )
        st.sidebar.markdown('<div style="margin-bottom: -10px; margin-top: -5px;"><hr /></div>', unsafe_allow_html=True)

        # Definir menus baseados no perfil do usuário
        if st.session_state.user_info['role'].lower() == 'admin':
            # Menu para administradores - reorganizado por categorias
            menu_categories = [
                "🕒 Gestão de Tempo", 
                "📊 Relatórios", 
                "📧 Comunicações", 
                "💼 Gestão", 
                "⚙️ Configurações"
            ]
            
            selected_category = st.sidebar.selectbox("Categoria", menu_categories, key="admin_category")
            
            if selected_category == "🕒 Gestão de Tempo":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["Registro de Horas", "Calendário"],
                    key="time_options"
                )
                
                if menu == "Registro de Horas":
                    timesheet_page()
                elif menu == "Calendário":
                    calendar_page()
                    
            elif selected_category == "📊 Relatórios":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["Projetos", "Produtividade por Equipe", 
                     "Produtividade por Usuário",
                     "Horas Trabalhadas", "Dashboard"],
                    key="report_options"
                )
                
                if menu == "Projetos":
                    reports_page()
                
                #elif menu == "Progresso de Fases":
                    #phase_progress_dashboard()
                elif menu == "Produtividade por Equipe":
                    team_productivity_page()
                elif menu == "Produtividade por Usuário":
                    user_productivity_page()
                elif menu == "Horas Trabalhadas":
                    worked_hours_report()
                elif menu == "Dashboard":
                    dashboard_page()
                
            elif selected_category == "📧 Comunicações":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["Email Indicadores Colaboradores", 
                     "Email Indicadores de Projeto", "Email Indicadores de Faturação", "Email Indicadores Comerciais",
                     "Alertas de Projetos"],
                    key="email_options"
                )
                
                if menu == "Email Indicadores Colaboradores":
                    collaborator_email_report()
                elif menu == "Email Indicadores de Projeto":
                    project_email_report()
                elif menu == "Email Indicadores de Faturação":
                    revenue_email_report()
                elif menu == "Email Indicadores Comerciais":
                    commercial_indicators_email()
                #elif menu == "Reuniões Comerciais":
                    #commercial_meetings_report()
                elif menu == "Alertas de Projetos":
                    project_status_email()
                
            elif selected_category == "💼 Gestão":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["CRM Comercial", "Faturação"],
                    key="mgmt_options"
                )
                
                if menu == "CRM Comercial":
                    from commercial_crm import crm_page
                    crm_page()
                elif menu == "Faturação":
                    billing_page()
                #elif menu == "Relatório Executivo":
                    #executive_dashboard_email()
                
            elif selected_category == "⚙️ Configurações":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["Gerais", "Metas Anuais"],
                    key="config_options"
                )
                
                if menu == "Gerais":
                    app_main()
                elif menu == "Metas Anuais":
                    from annual_targets import annual_targets_page
                    annual_targets_page()
                #elif menu == "Fases de Projetos":
                    #project_phases_page()

         # Menu para líderes de equipe
        elif st.session_state.user_info['role'].lower() == 'leader':
            menu_categories = [
                "🕒 Gestão de Tempo",
                "📊 Análise de Projetos"
            ]
            
            selected_category = st.sidebar.selectbox("Categoria", menu_categories, key="leader_category")
            
            if selected_category == "🕒 Gestão de Tempo":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["Registro de Horas", "Calendário"],
                    key="leader_time"
                )
                
                if menu == "Registro de Horas":
                    timesheet_page()
                elif menu == "Calendário":
                    calendar_page()
            
            elif selected_category == "📊 Análise de Projetos":
                menu = st.sidebar.selectbox(
                    "Opções",
                    ["Projetos", "Horas Trabalhadas"],
                    key="leader_analysis"
                )
                
                if menu == "Horas Trabalhadas":
                    worked_hours_report()
                #elif menu == "Produtividade da Equipe":
                    #team_productivity_page()
                #elif menu == "Ausências":
                    #ausencias_report_page()
                elif menu == "Projetos":
                    reports_page()
                    
        # Menu para usuários regulares
        else:
            menu = st.sidebar.selectbox(
                "Menu",
                ["Registro de Horas", "Calendário", "Horas Trabalhadas"],
                key="user_menu"
            )
            
            if menu == "Registro de Horas":
                timesheet_page()
            elif menu == "Calendário":
                calendar_page()
            elif menu == "Horas Trabalhadas":
                worked_hours_report()
            #elif menu == "Metas":
                #user_targets_dashboard()
                
        # Alterar senha (para todos exceto admin)
        if st.session_state.user_info['role'] != 'admin':
            if 'show_password_form' not in st.session_state:
                st.session_state.show_password_form = False

            if st.sidebar.button("🔑 Alterar Senha", type="secondary", use_container_width=True, key="alter_password"):
                st.session_state.show_password_form = True

            if st.session_state.show_password_form:
                st.title("Alteração de Senha")
                with st.form("form_alterar_senha", clear_on_submit=False):
                    senha_atual = st.text_input("Senha Atual", type="password")
                    nova_senha = st.text_input("Nova Senha", type="password")
                    confirmar_senha = st.text_input("Confirmar Nova Senha", type="password")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.form_submit_button("Confirmar"):
                            dados_usuario = UserManager().read(st.session_state.user_info['user_id'])
                            hash_atual = hashlib.sha256(str(senha_atual).encode()).hexdigest()

                            if hash_atual != dados_usuario.iloc[0]['password']:
                                st.error("Senha atual incorreta.")
                            elif nova_senha != confirmar_senha:
                                st.error("As senhas não coincidem.")
                            elif len(nova_senha) < 6:
                                st.error("A nova senha deve ter pelo menos 6 caracteres.")
                            else:
                                user_manager = UserManager()
                                nova_senha_hash = user_manager.hash_password(nova_senha)
                                user_manager.update(
                                    st.session_state.user_info['user_id'],
                                    {'password': nova_senha_hash}
                                )
                                st.success("Senha alterada com sucesso!")
                                st.session_state.show_password_form = False
                                time.sleep(1)
                                st.rerun()
                    
                    with col2:
                        if st.form_submit_button("Cancelar"):
                            st.session_state.show_password_form = False
                            st.rerun()

        # Botão de Logout e Manual do Utilizador (para todos)
        st.sidebar.markdown('<div style="margin-bottom: -15px; margin-top: -5px;"><hr /></div>', unsafe_allow_html=True)
        
        # Botões com menos padding
        col1, col2 = st.sidebar.columns(2)
        
        # Botão para download do manual
        with open("manual.pdf", "rb") as pdf_file:
            PDFbyte = pdf_file.read()

            col1.download_button(
                label="📖 Manual",
                data=PDFbyte,
                file_name="manual.pdf",
                mime='application/octet-stream',
                use_container_width=True
            )
        
        # Botão de logout com menos espaço
        if col2.button("↩️ Sair", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        st.sidebar.markdown(
            """
            <p style='text-align: center; color: #666666; font-size: 0.7em; margin-top: -5px;'>
            Versão 1.0
            </p>
            """, 
            unsafe_allow_html=True
        )


class UserManager:
    def __init__(self):
        self.table_name = 'utilizadores'  # Nome da tabela no banco de dados
    
    def load_data(self):
        # Substituir leitura do Excel por leitura do banco de dados
        try:
            return read_table_from_db(self.table_name)
        except Exception as e:
            # Em caso de erro (tabela não existe, etc.), retorna um DataFrame vazio
            print(f"Erro ao carregar dados: {e}")
            return pd.DataFrame()
    
    def save_data(self, df):
        # Substituir gravação em Excel por gravação no banco de dados
        save_table_to_db(df, self.table_name)
    
    def create(self, data):
        df = self.load_data()
        if df.empty:
            new_id = 1
        else:
            new_id = df['user_id'].max() + 1
        data['user_id'] = new_id
        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        self.save_data(df)
        return new_id
    
    def read(self, id=None):
        df = self.load_data()
        if id is None:
            return df
        return df[df['user_id'] == id]
    
    def update(self, id, data):
        df = self.load_data()
        for key, value in data.items():
            df.loc[df['user_id'] == id, key] = value
        self.save_data(df)
    
    def delete(self, id):
        df = self.load_data()
        df = df[df['user_id'] != id]
        self.save_data(df)

    def hash_password(self, password):
        return hashlib.sha256(str(password).encode()).hexdigest()
    
            
if __name__ == "__main__":
    main()