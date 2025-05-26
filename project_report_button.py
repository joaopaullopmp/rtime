# project_report_button.py
import streamlit as st
from project_report_export import download_project_report

def add_report_export_button(projeto_info, client_name, db_manager):
    """
    Adiciona um botão para exportar o relatório do projeto como PDF
    
    Args:
        projeto_info: Informações do projeto
        client_name: Nome do cliente
        db_manager: Instância do DatabaseManager
    """
    st.markdown("---")
    st.subheader("📄 Exportar Relatório")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.write("""
        Exporte um relatório detalhado deste projeto em formato PDF. 
        O relatório inclui todas as informações apresentadas nesta página, 
        formatadas de maneira profissional para impressão ou compartilhamento.
        """)
    
    with col2:
        # Adicionar o botão de download usando a função do módulo project_report_export
        download_project_report(projeto_info, client_name, db_manager)