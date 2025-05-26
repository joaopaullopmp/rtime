import streamlit as st
import logging
import pandas as pd
import io
import os
import smtplib
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
import calendar
from fpdf import FPDF
from database_manager import DatabaseManager
from annual_targets import AnnualTargetManager
from report_utils import calcular_dias_uteis_projeto
from risk_reports import calcular_risco_projeto

# Configuração do logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

def project_email_report():
    """
    Módulo para geração e envio de relatórios de indicadores de projetos por email
    """
    logging.info("Iniciando a geração do relatório de indicadores de projetos.")
    st.title("📧 Relatório de Indicadores de Projetos")
    
    # Inicializar gerenciadores
    db_manager = DatabaseManager()
    annual_target_manager = AnnualTargetManager()
    
    # Verificar se o usuário é administrador
    if st.session_state.user_info['role'].lower() != 'admin':
        st.warning("Esta funcionalidade é exclusiva para administradores.")
        return
    
    # Carregamento de dados básicos
    try:
        projects_df = db_manager.query_to_df("SELECT * FROM projects")
        clients_df = db_manager.query_to_df("SELECT * FROM clients")
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        rates_df = db_manager.query_to_df("SELECT * FROM rates")
        groups_df = db_manager.query_to_df("SELECT * FROM groups WHERE active = 1")
        
        # Converter datas
        projects_df['start_date'] = pd.to_datetime(projects_df['start_date'], format='mixed', errors='coerce')
        projects_df['end_date'] = pd.to_datetime(projects_df['end_date'], format='mixed', errors='coerce')
        timesheet_df['start_date'] = pd.to_datetime(timesheet_df['start_date'], format='mixed', errors='coerce')
        
        # Obter tipos de projetos únicos
        project_types = sorted(projects_df['project_type'].unique().tolist())
        
    except Exception as e:
        st.error(f"Erro ao carregar dados: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return
    
    # Interface para configuração do relatório
    st.subheader("Configuração do Relatório")
    
    with st.form("email_report_config"):
        # Seleção de destinatários
        recipients = st.text_input(
            "Destinatários (separados por vírgula)",
            help="Email dos destinatários separados por vírgula"
        )
        
        # Assunto do email
        subject = st.text_input(
            "Assunto do Email", 
            f"Relatório de Indicadores de Projetos - {datetime.now().strftime('%d/%m/%Y')}"
        )
        
        # Filtros para o relatório
        col1, col2 = st.columns(2)
        
        with col1:
            # Seleção de período
            report_period = st.selectbox(
                "Período do Relatório",
                ["Mês Atual", "Mês Anterior", "Últimos 3 Meses", "Ano Atual", "Período Personalizado"]
            )
            
            if report_period == "Período Personalizado":
                start_date = st.date_input(
                    "Data de Início",
                    value=datetime.now().replace(day=1)
                )
                end_date = st.date_input(
                    "Data de Término",
                    value=datetime.now()
                )
            else:
                # Definir datas automaticamente conforme o período selecionado
                today = datetime.now()
                
                if report_period == "Mês Atual":
                    start_date = today.replace(day=1)
                    last_day = calendar.monthrange(today.year, today.month)[1]
                    end_date = today.replace(day=last_day)
                    
                elif report_period == "Mês Anterior":
                    if today.month == 1:
                        previous_month = 12
                        previous_year = today.year - 1
                    else:
                        previous_month = today.month - 1
                        previous_year = today.year
                        
                    start_date = datetime(previous_year, previous_month, 1)
                    last_day = calendar.monthrange(previous_year, previous_month)[1]
                    end_date = datetime(previous_year, previous_month, last_day)
                    
                elif report_period == "Últimos 3 Meses":
                    # Calcular 3 meses atrás
                    if today.month <= 3:
                        months_back = today.month + 12 - 3
                        year_back = today.year - 1
                    else:
                        months_back = today.month - 3
                        year_back = today.year
                        
                    start_date = datetime(year_back, months_back, 1)
                    last_day = calendar.monthrange(today.year, today.month)[1]
                    end_date = today.replace(day=last_day)
                    
                elif report_period == "Ano Atual":
                    start_date = datetime(today.year, 1, 1)
                    end_date = datetime(today.year, 12, 31)
            
            # Seleção de tipos de projeto (múltipla)
            project_type_options = ["Todos"] + project_types
            selected_project_types = st.multiselect(
                "Tipos de Projeto",
                options=project_type_options,
                default=["Todos"]
            )
        
        with col2:
            # Seleção de equipes (múltipla)
            team_options = ["Todas"] + sorted(groups_df['group_name'].tolist())
            selected_teams = st.multiselect(
                "Equipes", 
                options=team_options,
                default=["Todas"]
            )
            
            # Seleção de clientes (múltipla)
            client_options = ["Todos"] + sorted(clients_df['name'].tolist())
            selected_clients = st.multiselect(
                "Clientes", 
                options=client_options,
                default=["Todos"]
            )
            
            # Formato do relatório
            report_format = st.radio(
                "Formato do Relatório",
                ["PDF", "Excel", "PDF e Excel"]
            )
            
            # Opções adicionais
            show_financial = st.checkbox("Incluir Informações Financeiras", value=True, 
                                       help="Incluir custos e indicadores financeiros")
            show_hour_details = st.checkbox("Incluir Detalhes de Horas", value=True,
                                          help="Incluir detalhamento de horas por projeto")
        
        # Conteúdo do email
        email_message = st.text_area(
            "Mensagem do Email",
            """Prezados,

Em anexo, relatório de indicadores de desempenho dos projetos.

Atenciosamente,
Equipe de Gestão"""
        )
        
        # SMTP settings (collapsed by default)
        with st.expander("Configurações de SMTP"):
            smtp_server = st.text_input("Servidor SMTP", "smtp.office365.com")
            smtp_port = st.number_input("Porta SMTP", value=587, step=1)
            smtp_user = st.text_input("Usuário SMTP", "notifications@grupoerre.pt")
            smtp_password = st.text_input("Senha SMTP", type="password", value="9FWkMpK8tif2lY4")
            use_tls = st.checkbox("Usar TLS", value=True)
        
        # Botão para gerar e enviar o relatório
        submit_button = st.form_submit_button("Gerar e Enviar Relatório")
    
    if submit_button:
        # Validar entradas
        if not recipients:
            st.error("Por favor, informe pelo menos um destinatário.")
            return
        
        with st.spinner("Gerando e enviando relatório..."):
            # Criar diretório temporário para os arquivos
            temp_dir = tempfile.mkdtemp()
            
            # Gerar relatórios conforme formato selecionado
            pdf_path = None
            excel_path = None
            
            if "PDF" in report_format:
                pdf_path = os.path.join(temp_dir, "relatorio_indicadores_projetos.pdf")
                pdf_result = generate_project_pdf_report(
                    pdf_path,
                    db_manager,
                    annual_target_manager,
                    start_date,
                    end_date,
                    selected_teams,
                    selected_clients,
                    selected_project_types,
                    show_financial,
                    show_hour_details
                )
                
                # Verificar se o PDF foi gerado com sucesso
                if not pdf_result or not os.path.exists(pdf_path):
                    st.error(f"Falha ao gerar o PDF. Verifique os logs para mais detalhes.")
                    pdf_path = None

            if "Excel" in report_format:
                excel_path = os.path.join(temp_dir, "relatorio_indicadores_projetos.xlsx")
                try:
                    # Obter dados dos projetos para o Excel
                    projetos = get_project_indicators(
                        db_manager,
                        annual_target_manager,
                        start_date,
                        end_date,
                        selected_teams,
                        selected_clients,
                        selected_project_types
                    )
                    
                    excel_result = generate_project_excel_report(
                        excel_path,
                        projetos,
                        start_date,
                        end_date,
                        selected_teams,
                        selected_clients,
                        selected_project_types,
                        show_financial,
                        show_hour_details
                    )
                    
                    if excel_result and os.path.exists(excel_path):
                        st.success("Excel gerado com sucesso!")
                    else:
                        st.error("Falha ao gerar o Excel.")
                        excel_path = None
                except Exception as e:
                    st.error(f"Erro ao gerar Excel: {str(e)}")
                    excel_path = None
            
            # Enviar email com os relatórios gerados se pelo menos um foi gerado com sucesso
            if pdf_path or excel_path:
                email_success = send_email(
                    recipients.split(','), 
                    subject, 
                    email_message, 
                    pdf_path, 
                    excel_path, 
                    smtp_server, 
                    smtp_port, 
                    smtp_user, 
                    smtp_password, 
                    use_tls
                )
                
                # Botões para download local dos relatórios
                st.markdown("### Download dos Relatórios")
                
                if pdf_path and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                        st.download_button(
                            label="📥 Baixar PDF",
                            data=pdf_bytes,
                            file_name="relatorio_indicadores_projetos.pdf",
                            mime="application/pdf"
                        )
                
                if excel_path and os.path.exists(excel_path):
                    with open(excel_path, "rb") as excel_file:
                        excel_bytes = excel_file.read()
                        st.download_button(
                            label="📥 Baixar Excel",
                            data=excel_bytes,
                            file_name="relatorio_indicadores_projetos.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.error("Não foi possível gerar nenhum dos relatórios. Verifique os logs para mais detalhes.")


def generate_project_pdf_report(
    output_path,
    db_manager,
    annual_target_manager,
    start_date,
    end_date,
    selected_teams,
    selected_clients,
    selected_project_types,
    show_financial=True,
    show_hour_details=True
):
    """
    Gera relatório PDF com indicadores de projetos
    """
    logging.info("Iniciando a geração do relatório PDF de projetos.")

    try:
        # Inicializar PDF
        class PDF(FPDF):
            def header(self):
                # Logo apenas a partir da segunda página
                if self.page_no() != 1:  # Não mostrar o logo na primeira página (capa)
                    if os.path.exists('logo.png'):
                        self.image('logo.png', 10, 8, 33)
                    
                    # Título
                    self.set_font('Arial', 'B', 15)
                    self.cell(80)
                    self.cell(30, 25, 'Indicadores de Performance de Projetos', 0, 0, 'C')
                    self.ln(20)
            
            def footer(self):
                # Posicionar a 1.5 cm do final
                self.set_y(-15)
                # Fonte Arial itálico 8
                self.set_font('Arial', 'I', 8)
                # Número da página
                self.cell(0, 10, f'Pagina {self.page_no()}/{{nb}}', 0, 0, 'C')
        
        pdf = PDF()
        pdf.alias_nb_pages()
        
        # Capa do relatório
        pdf.add_page()

        # Adicionar o logo apenas uma vez na capa
        if os.path.exists('logo.png'):
            pdf.image('logo.png', x=10, y=20, w=60)

        pdf.set_font('Arial', 'B', 20)
        pdf.set_xy(0, 100)
        pdf.cell(210, 20, 'Relatorio - Performance de Projetos', 0, 1, 'C')
        
        # Período do relatório
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 15, f"Periodo: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}", 0, 1, 'C')
        
        # Filtros aplicados
        pdf.set_font('Arial', '', 12)
        pdf.ln(10)
        
        teams_str = ', '.join(selected_teams) if "Todas" not in selected_teams else "Todas as equipes"
        clients_str = ', '.join(selected_clients) if "Todos" not in selected_clients else "Todos os clientes"
        types_str = ', '.join(selected_project_types) if "Todos" not in selected_project_types else "Todos os tipos"
        
        pdf.set_x(40)
        pdf.cell(0, 8, f"Equipes: {teams_str}", 0, 1)
        pdf.set_x(40)
        pdf.cell(0, 8, f"Clientes: {clients_str}", 0, 1)
        pdf.set_x(40)
        pdf.cell(0, 8, f"Tipos de Projeto: {types_str}", 0, 1)
        
        # Data de geração
        pdf.set_font('Arial', 'I', 10)
        pdf.set_y(-30)
        pdf.cell(0, 10, f"Gerado automaticamente em {datetime.now().strftime('%d/%m/%Y as %H:%M')}", 0, 1, 'C')
        
        # Obter indicadores de projetos
        projetos = get_project_indicators(
            db_manager,
            annual_target_manager,
            start_date,
            end_date,
            selected_teams,
            selected_clients,
            selected_project_types
        )
        
        if not projetos:
            # Página de erro
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Sem Dados Disponíveis', 0, 1, 'C')
            
            pdf.set_font('Arial', '', 12)
            pdf.ln(10)
            pdf.multi_cell(0, 8, "Não foram encontrados dados de projetos para o período e filtros selecionados.")
            
            # Salvar o PDF mesmo assim
            pdf.output(output_path)
            return True
            
        # Converter para DataFrame
        df = pd.DataFrame(projetos)
        
        # Resumo executivo
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Resumo Executivo', 0, 1)
        
        # Calcular dias úteis do período
        dias_uteis = calcular_dias_uteis_projeto(start_date.date(), end_date.date())
        
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 8, f"Este relatório apresenta uma visão consolidada dos indicadores de desempenho dos projetos no período de {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}, cobrindo {dias_uteis} dias úteis.")
        
        # Dados resumidos para o sumário
        total_projetos = len(df)
        projetos_alto_risco = sum(df['risk_level'] == 'Alto')
        projetos_medio_risco = sum(df['risk_level'] == 'Médio')
        projetos_baixo_risco = sum(df['risk_level'] == 'Baixo')
        percentual_horas_media = df['hours_percentage'].mean() if 'hours_percentage' in df.columns else 0
        
        if show_financial and 'cost_percentage' in df.columns:
            percentual_custo_media = df['cost_percentage'].mean()
            cpi_medio = df['cpi'].mean()
        
        # Adicionar tabela de resumo
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, "Resumo dos Indicadores:", 0, 1)
        
        # Cabeçalho da tabela
        pdf.set_fill_color(31, 119, 180)  # Azul
        pdf.set_text_color(255, 255, 255)  # Branco
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(100, 8, "Métrica", 1, 0, 'L', True)
        pdf.cell(35, 8, "Valor", 1, 1, 'C', True)
        
        # Dados da tabela com cores alternadas
        pdf.set_text_color(0, 0, 0)  # Preto
        
        # Total de projetos
        pdf.set_font('Arial', '', 10)
        pdf.cell(100, 8, "Total de Projetos", 1, 0, 'L')
        pdf.cell(35, 8, str(total_projetos), 1, 1, 'C')
        
        # Distribuição de risco
        pdf.set_fill_color(245, 245, 245)  # Cinza claro
        pdf.cell(100, 8, "Projetos em Alto Risco", 1, 0, 'L', True)
        pdf.cell(35, 8, f"{projetos_alto_risco} ({(projetos_alto_risco/total_projetos*100):.1f}%)", 1, 1, 'C', True)
        
        pdf.cell(100, 8, "Projetos em Médio Risco", 1, 0, 'L')
        pdf.cell(35, 8, f"{projetos_medio_risco} ({(projetos_medio_risco/total_projetos*100):.1f}%)", 1, 1, 'C')
        
        pdf.set_fill_color(245, 245, 245)  # Cinza claro
        pdf.cell(100, 8, "Projetos em Baixo Risco", 1, 0, 'L', True)
        pdf.cell(35, 8, f"{projetos_baixo_risco} ({(projetos_baixo_risco/total_projetos*100):.1f}%)", 1, 1, 'C', True)
        
        # Percentual médio de horas gastas
        pdf.cell(100, 8, "Percentual Médio de Horas Gastas", 1, 0, 'L')
        pdf.cell(35, 8, f"{percentual_horas_media:.1f}%", 1, 1, 'C')
        
        # Informações financeiras (se solicitado)
        if show_financial and 'cost_percentage' in df.columns and 'cpi' in df.columns:
            pdf.set_fill_color(245, 245, 245)  # Cinza claro
            pdf.cell(100, 8, "Percentual Médio de Custo Gasto", 1, 0, 'L', True)
            pdf.cell(35, 8, f"{percentual_custo_media:.1f}%", 1, 1, 'C', True)
            
            pdf.cell(100, 8, "CPI Médio", 1, 0, 'L')
            pdf.cell(35, 8, f"{cpi_medio:.2f}", 1, 1, 'C')

        # Projetos em alto risco
        if projetos_alto_risco > 0:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 14)
            pdf.set_fill_color(255, 230, 230)  # Vermelho claro
            pdf.cell(0, 10, "Projetos em Alto Risco", 0, 1, 'L', True)
            
            # Listar projetos em alto risco
            pdf.ln(5)
            
            high_risk_projects = df[df['risk_level'] == 'Alto'].sort_values('cpi')
            
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 9)
            
            # Colunas da tabela
            if show_financial:
                pdf.cell(45, 8, "Projeto", 1, 0, 'L', True)
                pdf.cell(35, 8, "Cliente", 1, 0, 'L', True)
                pdf.cell(25, 8, "% Horas", 1, 0, 'C', True)
                pdf.cell(25, 8, "% Custo", 1, 0, 'C', True)
                pdf.cell(25, 8, "CPI", 1, 0, 'C', True)
                pdf.cell(35, 8, "Razão", 1, 1, 'L', True)
            else:
                pdf.cell(60, 8, "Projeto", 1, 0, 'L', True)
                pdf.cell(40, 8, "Cliente", 1, 0, 'L', True)
                pdf.cell(30, 8, "% Horas", 1, 0, 'C', True)
                pdf.cell(60, 8, "Razão", 1, 1, 'L', True)
            
            # Dados da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 8)
            
            for idx, row in high_risk_projects.iterrows():
                # Alternar cores de fundo
                if idx % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)  # Branco
                    fill = False
                else:
                    pdf.set_fill_color(245, 245, 245)  # Cinza claro
                    fill = True
                    
                # Verificar espaço disponível
                if pdf.get_y() > 250:  # Se estiver próximo ao fim da página
                    pdf.add_page()
                    
                    # Repetir o cabeçalho
                    pdf.set_fill_color(31, 119, 180)  # Azul
                    pdf.set_text_color(255, 255, 255)  # Branco
                    pdf.set_font('Arial', 'B', 9)
                    
                    if show_financial:
                        pdf.cell(45, 8, "Projeto", 1, 0, 'L', True)
                        pdf.cell(35, 8, "Cliente", 1, 0, 'L', True)
                        pdf.cell(25, 8, "% Horas", 1, 0, 'C', True)
                        pdf.cell(25, 8, "% Custo", 1, 0, 'C', True)
                        pdf.cell(25, 8, "CPI", 1, 0, 'C', True)
                        pdf.cell(35, 8, "Razão", 1, 1, 'L', True)
                    else:
                        pdf.cell(60, 8, "Projeto", 1, 0, 'L', True)
                        pdf.cell(40, 8, "Cliente", 1, 0, 'L', True)
                        pdf.cell(30, 8, "% Horas", 1, 0, 'C', True)
                        pdf.cell(60, 8, "Razão", 1, 1, 'L', True)
                    
                    pdf.set_text_color(0, 0, 0)  # Preto
                    pdf.set_font('Arial', '', 8)
                
                # Truncar razão para evitar textos muito longos
                razao = row['risk_reason']
                razao_truncada = (razao[:50] + '...') if len(razao) > 50 else razao
                
                if show_financial:
                    pdf.cell(45, 8, row['project_name'], 1, 0, 'L', fill)
                    pdf.cell(35, 8, row['client_name'], 1, 0, 'L', fill)
                    pdf.cell(25, 8, f"{row['hours_percentage']:.1f}%", 1, 0, 'C', fill)
                    pdf.cell(25, 8, f"{row['cost_percentage']:.1f}%", 1, 0, 'C', fill)
                    pdf.cell(25, 8, f"{row['cpi']:.2f}", 1, 0, 'C', fill)
                    pdf.cell(35, 8, razao_truncada, 1, 1, 'L', fill)
                else:
                    pdf.cell(60, 8, row['project_name'], 1, 0, 'L', fill)
                    pdf.cell(40, 8, row['client_name'], 1, 0, 'L', fill)
                    pdf.cell(30, 8, f"{row['hours_percentage']:.1f}%", 1, 0, 'C', fill)
                    pdf.cell(60, 8, razao_truncada, 1, 1, 'L', fill)

        # Tabela completa de projetos
        pdf.add_page()
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, "Tabela Completa de Indicadores de Projetos", 0, 1)
        
        # Ordenar alfabeticamente por cliente, depois por projeto
        df_sorted = df.sort_values(['client_name', 'project_name'])
        
        # Cabeçalho da tabela
        pdf.set_fill_color(31, 119, 180)  # Azul
        pdf.set_text_color(255, 255, 255)  # Branco
        pdf.set_font('Arial', 'B', 8)
        
        # Definir larguras das colunas com base nas informações financeiras
        if show_financial:
            col_widths = [40, 35, 25, 15, 15, 15, 15, 15, 15]
            pdf.cell(col_widths[0], 8, "Projeto", 1, 0, 'L', True)
            pdf.cell(col_widths[1], 8, "Cliente", 1, 0, 'L', True)
            pdf.cell(col_widths[2], 8, "Tipo", 1, 0, 'L', True)
            pdf.cell(col_widths[3], 8, "Horas Atual", 1, 0, 'C', True)
            pdf.cell(col_widths[4], 8, "Horas Total", 1, 0, 'C', True)
            pdf.cell(col_widths[5], 8, "% Horas", 1, 0, 'C', True)
            pdf.cell(col_widths[6], 8, "% Custo", 1, 0, 'C', True)
            pdf.cell(col_widths[7], 8, "CPI", 1, 0, 'C', True)
            pdf.cell(col_widths[8], 8, "Risco", 1, 1, 'C', True)
        else:
            col_widths = [50, 40, 30, 20, 20, 20, 20]
            pdf.cell(col_widths[0], 8, "Projeto", 1, 0, 'L', True)
            pdf.cell(col_widths[1], 8, "Cliente", 1, 0, 'L', True)
            pdf.cell(col_widths[2], 8, "Tipo", 1, 0, 'L', True)
            pdf.cell(col_widths[3], 8, "Horas Atual", 1, 0, 'C', True)
            pdf.cell(col_widths[4], 8, "Horas Total", 1, 0, 'C', True)
            pdf.cell(col_widths[5], 8, "% Horas", 1, 0, 'C', True)
            pdf.cell(col_widths[6], 8, "Risco", 1, 1, 'C', True)
            
        # Dados da tabela
        pdf.set_text_color(0, 0, 0)  # Preto
        pdf.set_font('Arial', '', 7)
        
        # Contador para alternar cores
        row_count = 0
        
        for _, row in df_sorted.iterrows():
            # Verificar espaço disponível para nova página
            if pdf.get_y() > 250:  # Se estiver próximo ao fim da página
                pdf.add_page()
                
                # Repetir o cabeçalho
                pdf.set_fill_color(31, 119, 180)  # Azul
                pdf.set_text_color(255, 255, 255)  # Branco
                pdf.set_font('Arial', 'B', 8)
                
                if show_financial:
                    pdf.cell(col_widths[0], 8, "Projeto", 1, 0, 'L', True)
                    pdf.cell(col_widths[1], 8, "Cliente", 1, 0, 'L', True)
                    pdf.cell(col_widths[2], 8, "Tipo", 1, 0, 'L', True)
                    pdf.cell(col_widths[3], 8, "Horas Atual", 1, 0, 'C', True)
                    pdf.cell(col_widths[4], 8, "Horas Total", 1, 0, 'C', True)
                    pdf.cell(col_widths[5], 8, "% Horas", 1, 0, 'C', True)
                    pdf.cell(col_widths[6], 8, "% Custo", 1, 0, 'C', True)
                    pdf.cell(col_widths[7], 8, "CPI", 1, 0, 'C', True)
                    pdf.cell(col_widths[8], 8, "Risco", 1, 1, 'C', True)
                else:
                    pdf.cell(col_widths[0], 8, "Projeto", 1, 0, 'L', True)
                    pdf.cell(col_widths[1], 8, "Cliente", 1, 0, 'L', True)
                    pdf.cell(col_widths[2], 8, "Tipo", 1, 0, 'L', True)
                    pdf.cell(col_widths[3], 8, "Horas Atual", 1, 0, 'C', True)
                    pdf.cell(col_widths[4], 8, "Horas Total", 1, 0, 'C', True)
                    pdf.cell(col_widths[5], 8, "% Horas", 1, 0, 'C', True)
                    pdf.cell(col_widths[6], 8, "Risco", 1, 1, 'C', True)
                
                pdf.set_text_color(0, 0, 0)  # Preto
                pdf.set_font('Arial', '', 7)
            
            # Alternar cores de fundo
            if row_count % 2 == 0:
                pdf.set_fill_color(255, 255, 255)  # Branco
                fill = False
            else:
                pdf.set_fill_color(245, 245, 245)  # Cinza claro
                fill = True
            
            # Verificar se é projeto de alto risco e destacar
            if row['risk_level'] == 'Alto':
                pdf.set_text_color(255, 0, 0)  # Texto vermelho para alto risco
            
            # Truncar nomes longos
            project_name = row['project_name']
            if len(project_name) > 50:
                project_name = project_name[:49] + '…'
                
            client_name = row['client_name']
            if len(client_name) > 15:
                client_name = client_name[:14] + '…'
                
            project_type = row['project_type']
            if len(project_type) > 12:
                project_type = project_type[:11] + '…'
            
            # Preencher os dados
            if show_financial:
                pdf.cell(col_widths[0], 7, project_name, 1, 0, 'L', fill)
                pdf.cell(col_widths[1], 7, client_name, 1, 0, 'L', fill)
                pdf.cell(col_widths[2], 7, project_type, 1, 0, 'L', fill)
                pdf.cell(col_widths[3], 7, f"{row['realized_hours']:.1f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[4], 7, f"{row['total_hours']:.1f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[5], 7, f"{row['hours_percentage']:.1f}%", 1, 0, 'C', fill)
                pdf.cell(col_widths[6], 7, f"{row['cost_percentage']:.1f}%", 1, 0, 'C', fill)
                pdf.cell(col_widths[7], 7, f"{row['cpi']:.2f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[8], 7, row['risk_level'], 1, 1, 'C', fill)
            else:
                pdf.cell(col_widths[0], 7, project_name, 1, 0, 'L', fill)
                pdf.cell(col_widths[1], 7, client_name, 1, 0, 'L', fill)
                pdf.cell(col_widths[2], 7, project_type, 1, 0, 'L', fill)
                pdf.cell(col_widths[3], 7, f"{row['realized_hours']:.1f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[4], 7, f"{row['total_hours']:.1f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[5], 7, f"{row['hours_percentage']:.1f}%", 1, 0, 'C', fill)
                pdf.cell(col_widths[6], 7, row['risk_level'], 1, 1, 'C', fill)
            
            # Restaurar cor normal
            pdf.set_text_color(0, 0, 0)
            
            # Incrementar contador de linha
            row_count += 1
        
        # Legenda de cores de risco
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, "Legenda de Risco:", 0, 1)
        
        pdf.set_font('Arial', '', 9)
        
        # Alto risco
        pdf.set_text_color(255, 0, 0)  # Vermelho
        pdf.cell(20, 8, "Alto:", 0, 0)
        pdf.set_text_color(0, 0, 0)  # Preto
        pdf.cell(170, 8, "CPI < 0.9 - O projeto está acima do orçamento previsto", 0, 1)
        
        # Médio risco
        pdf.set_fill_color(255, 193, 7)  # Amarelo
        pdf.rect(12, pdf.get_y() + 3, 5, 5, 'F')
        pdf.cell(20, 8, "Médio:", 0, 0)
        pdf.cell(170, 8, "0.9 <= CPI < 1.1 - O projeto está próximo ao orçamento previsto", 0, 1)
        
        # Baixo risco
        pdf.set_fill_color(76, 175, 80)  # Verde
        pdf.rect(12, pdf.get_y() + 3, 5, 5, 'F')
        pdf.cell(20, 8, "Baixo:", 0, 0)
        pdf.cell(170, 8, "CPI >= 1.1 - O projeto está abaixo do orçamento previsto", 0, 1)
        
        # Explicação do CPI
        pdf.ln(5)
        pdf.set_font('Arial', 'I', 8)
        pdf.multi_cell(0, 5, "CPI (Cost Performance Index): Razão entre o valor planejado e o custo real. Um CPI maior que 1 indica que o projeto está gastando menos do que o planejado (positivo), enquanto um CPI menor que 1 indica que está gastando mais do que o planejado (negativo).")
        
        # Detalhes de horas por projeto (se solicitado)
        if show_hour_details and len(df) > 0:
            # Obter detalhes de horas por projeto
            # Criar uma tabela por projeto com informações de recursos utilizados
            for _, project_row in df.iterrows():
                project_id = project_row['project_id']
                project_name = project_row['project_name']
                
                # Obter entradas de timesheet para este projeto
                project_timesheet = timesheet_df[timesheet_df['project_id'] == project_id]
                
                if not project_timesheet.empty:
                    # Verificar se precisamos de uma nova página
                    if pdf.get_y() > 240:
                        pdf.add_page()
                    else:
                        pdf.ln(10)
                    
                    # Título do projeto
                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_fill_color(220, 220, 220)  # Cinza claro
                    pdf.cell(0, 10, f"Detalhamento de Horas - {project_name}", 0, 1, 'L', True)
                    
                    # Agrupar entradas por usuário
                    user_hours = {}
                    for _, entry in project_timesheet.iterrows():
                        user_id = entry['user_id']
                        hours = float(entry['hours'])
                        
                        if user_id not in user_hours:
                            user_hours[user_id] = {'regular': 0, 'extra': 0}
                        
                        # Verificar se é hora extra
                        is_overtime = entry.get('overtime', False)
                        if isinstance(is_overtime, (int, float)):
                            is_overtime = bool(is_overtime)
                        elif isinstance(is_overtime, str):
                            is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                        
                        if is_overtime:
                            user_hours[user_id]['extra'] += hours
                        else:
                            user_hours[user_id]['regular'] += hours
                    
                    # Total de horas do projeto
                    total_hours = sum([(u['regular'] + u['extra']) for u in user_hours.values()])
                    
                    # Criar tabela
                    pdf.set_font('Arial', 'B', 9)
                    pdf.set_fill_color(31, 119, 180)  # Azul
                    pdf.set_text_color(255, 255, 255)  # Branco
                    
                    # Cabeçalho da tabela
                    pdf.cell(60, 8, "Colaborador", 1, 0, 'L', True)
                    pdf.cell(30, 8, "Horas Normais", 1, 0, 'C', True)
                    pdf.cell(30, 8, "Horas Extras", 1, 0, 'C', True)
                    pdf.cell(30, 8, "Total", 1, 0, 'C', True)
                    pdf.cell(30, 8, "% do Projeto", 1, 1, 'C', True)
                    
                    # Linhas da tabela
                    pdf.set_font('Arial', '', 8)
                    pdf.set_text_color(0, 0, 0)  # Preto
                    
                    row_count = 0
                    for user_id, hours_data in user_hours.items():
                        # Obter nome do usuário
                        user_info = users_df[users_df['user_id'] == user_id]
                        if not user_info.empty:
                            user_name = f"{user_info['First_Name'].iloc[0]} {user_info['Last_Name'].iloc[0]}"
                        else:
                            user_name = f"Usuário ID: {user_id}"
                        
                        # Calcular total e percentual
                        user_total = hours_data['regular'] + hours_data['extra']
                        percentage = (user_total / total_hours * 100) if total_hours > 0 else 0
                        
                        # Alternar cores
                        if row_count % 2 == 0:
                            pdf.set_fill_color(255, 255, 255)  # Branco
                            fill = False
                        else:
                            pdf.set_fill_color(245, 245, 245)  # Cinza claro
                            fill = True
                        
                        # Dados
                        pdf.cell(60, 7, user_name, 1, 0, 'L', fill)
                        pdf.cell(30, 7, f"{hours_data['regular']:.1f}h", 1, 0, 'C', fill)
                        pdf.cell(30, 7, f"{hours_data['extra']:.1f}h", 1, 0, 'C', fill)
                        pdf.cell(30, 7, f"{user_total:.1f}h", 1, 0, 'C', fill)
                        pdf.cell(30, 7, f"{percentage:.1f}%", 1, 1, 'C', fill)
                        
                        row_count += 1
        
        # Salvar o PDF
        pdf.output(output_path)
        logging.info("PDF gerado com sucesso.")
        
        return True
    
    except Exception as e:
        import traceback
        error_msg = f"Erro ao gerar PDF: {str(e)}"
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        return False
    
def get_project_indicators(
    db_manager,
    annual_target_manager,
    start_date,
    end_date,
    selected_teams,
    selected_clients,
    selected_project_types
):
    """
    Obter indicadores de projetos para o relatório
    """
    try:
        # Carregar dados necessários
        projects_df = db_manager.query_to_df("SELECT * FROM projects")
        clients_df = db_manager.query_to_df("SELECT * FROM clients")
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        rates_df = db_manager.query_to_df("SELECT * FROM rates")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        
        # Converter datas
        projects_df['start_date'] = pd.to_datetime(projects_df['start_date'], format='mixed', errors='coerce')
        projects_df['end_date'] = pd.to_datetime(projects_df['end_date'], format='mixed', errors='coerce')
        
        # Filtragens
        filtered_projects = projects_df.copy()
        
        # Filtrar por status (apenas ativos) - ESTA É A LINHA MODIFICADA
        filtered_projects = filtered_projects[
            filtered_projects["status"].astype(str).str.strip().str.lower() == "active"
        ]
        
        # Filtrar por equipe
        if "Todas" not in selected_teams:
            # Obter IDs das equipes selecionadas
            team_ids = groups_df[groups_df["group_name"].isin(selected_teams)]['id'].tolist()
            
            # Filtrar projetos por equipe
            filtered_projects = filtered_projects[filtered_projects["group_id"].isin(team_ids)]
        
        # Filtrar por cliente
        if "Todos" not in selected_clients:
            # Obter IDs dos clientes selecionados
            client_ids = clients_df[clients_df["name"].isin(selected_clients)]['client_id'].tolist()
            
            # Filtrar projetos por cliente
            filtered_projects = filtered_projects[filtered_projects["client_id"].isin(client_ids)]
        
        # Filtrar por tipo de projeto
        if "Todos" not in selected_project_types:
            filtered_projects = filtered_projects[filtered_projects["project_type"].isin(selected_project_types)]
        
        if filtered_projects.empty:
            return []
        
        # Função para calcular dias úteis entre duas datas
        def calcular_dias_uteis(data_inicio, data_fim):
            """Calcula os dias úteis entre duas datas"""
            dias_uteis = 0
            data_atual = data_inicio
            
            while data_atual <= data_fim:
                if data_atual.weekday() < 5:  # 0-4 são dias úteis (seg-sex)
                    dias_uteis += 1
                data_atual += timedelta(days=1)
            
            return dias_uteis
        
        # Para cada projeto, calcular métricas
        project_indicators = []
        
        # Definir o início e fim do período para filtrar
        inicio_periodo = start_date
        fim_periodo = end_date
        
        # Converter datas em timesheet para datetime para facilitar filtros
        if 'start_date' in timesheet_df.columns:
            try:
                timesheet_df['start_date_dt'] = pd.to_datetime(timesheet_df['start_date'], format='mixed')
            except Exception as e:
                logging.warning(f"Erro ao converter datas: {e}")
                timesheet_df['start_date_dt'] = pd.to_datetime('2000-01-01')
        
        for _, project in filtered_projects.iterrows():
            try:
                # Obter registros de timesheet para o projeto
                project_entries = timesheet_df[timesheet_df["project_id"] == project["project_id"]]
                
                # Filtrar entradas do timesheet para o período selecionado
                period_entries = project_entries[
                    (project_entries["start_date_dt"] >= inicio_periodo) &
                    (project_entries["start_date_dt"] <= fim_periodo)
                ] if 'start_date_dt' in project_entries.columns else pd.DataFrame()
                
                # Verificar se existem dados migrados
                horas_migradas = 0
                if 'horas_realizadas_mig' in project and not pd.isna(project['horas_realizadas_mig']):
                    horas_migradas = float(project['horas_realizadas_mig'])
                
                custo_migrado = 0
                if 'custo_realizado_mig' in project and not pd.isna(project['custo_realizado_mig']):
                    custo_migrado = float(project['custo_realizado_mig'])
                
                # Calcular dias úteis totais do projeto
                dias_uteis_projeto = calcular_dias_uteis(project['start_date'], project['end_date'])
                
                # Calcular horas regulares e extras para o período
                hours_regular = 0
                hours_extra = 0
                
                if not period_entries.empty:
                    hours_regular = period_entries[~period_entries['overtime'].astype(bool)]['hours'].sum()
                    # Para horas extras, multiplicamos por 2 para contabilização
                    hours_extra_original = period_entries[period_entries['overtime'].astype(bool)]['hours'].sum()
                    hours_extra = hours_extra_original * 2
                
                # Total de horas no período (regulares + extras*2)
                period_hours = float(hours_regular + hours_extra)
                
                # Calcular custo realizado para o período
                period_cost = 0
                if not period_entries.empty:
                    for _, entry in period_entries.iterrows():
                        try:
                            user_id = entry['user_id']
                            hours = float(entry['hours'])
                            is_overtime = entry.get('overtime', False)
                            
                            # Converter para booleano
                            if isinstance(is_overtime, (int, float)):
                                is_overtime = bool(is_overtime)
                            elif isinstance(is_overtime, str):
                                is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                            
                            # Obter rate para o usuário
                            rate_value = None
                            if 'rate_value' in entry and not pd.isna(entry['rate_value']):
                                rate_value = float(entry['rate_value'])
                            else:
                                user_info = users_df[users_df['user_id'] == user_id]
                                
                                if not user_info.empty and not pd.isna(user_info['rate_id'].iloc[0]):
                                    rate_id = user_info['rate_id'].iloc[0]
                                    rate_info = rates_df[rates_df['rate_id'] == rate_id]
                                    
                                    if not rate_info.empty:
                                        rate_value = float(rate_info['rate_cost'].iloc[0])
                            
                            # Calcular custo com base no rate obtido
                            if rate_value:
                                entry_cost = hours * rate_value
                                
                                # Se for hora extra, multiplicar por 2
                                if is_overtime:
                                    period_cost += entry_cost * 2  # Dobro para horas extras
                                else:
                                    period_cost += entry_cost  # Normal para horas regulares
                        except Exception as e:
                            logging.warning(f"Erro ao processar entrada: {e}")
                
                # Calcular o percentual de tempo decorrido do projeto até a data final do período
                # Converter para datas no formato date()
                data_inicio = project['start_date'].date()
                data_fim = project['end_date'].date()
                data_atual = min(datetime.now().date(), data_fim)

                # Se o período final do relatório for anterior à data de início do projeto, 
                # não há progresso
                if end_date.date() <= data_inicio:
                    dias_uteis_totais = 0
                    dias_uteis_decorridos = 0
                    time_percentage = 0
                else:
                    # Limitar a data atual à data final do relatório, se necessário
                    data_atual = min(data_atual, end_date.date())
                    
                    # Calcular dias úteis totais do projeto
                    dias_uteis_totais = calcular_dias_uteis_projeto(data_inicio, data_fim)
                    
                    # Calcular dias úteis decorridos até a data atual
                    dias_uteis_decorridos = calcular_dias_uteis_projeto(data_inicio, data_atual)
                    
                    # Calcular percentual do tempo decorrido
                    time_percentage = (dias_uteis_decorridos / dias_uteis_totais * 100) if dias_uteis_totais > 0 else 0
                
                # Obter horas e custo totais planejados
                total_hours = float(project['total_hours']) if pd.notna(project['total_hours']) else 0
                total_cost = float(project['total_cost']) if pd.notna(project['total_cost']) else 0
                
                # Calcular horas e custo realizados totais (incluindo dados migrados)
                realized_hours = 0
                realized_cost = 0
                
                if not project_entries.empty:
                    # Calcular horas realizadas
                    regular_hours = project_entries[~project_entries['overtime'].astype(bool)]['hours'].sum()
                    extra_hours_original = project_entries[project_entries['overtime'].astype(bool)]['hours'].sum()
                    extra_hours = extra_hours_original * 2
                    realized_hours = regular_hours + extra_hours
                    
                    # Calcular custo realizado
                    for _, entry in project_entries.iterrows():
                        try:
                            user_id = entry['user_id']
                            hours = float(entry['hours'])
                            is_overtime = entry.get('overtime', False)
                            
                            # Converter para booleano
                            if isinstance(is_overtime, (int, float)):
                                is_overtime = bool(is_overtime)
                            elif isinstance(is_overtime, str):
                                is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                            
                            # Obter rate para o usuário
                            user_info = users_df[users_df['user_id'] == user_id]
                            if not user_info.empty and not pd.isna(user_info['rate_id'].iloc[0]):
                                rate_id = user_info['rate_id'].iloc[0]
                                rate_info = rates_df[rates_df['rate_id'] == rate_id]
                                
                                if not rate_info.empty:
                                    rate_value = float(rate_info['rate_cost'].iloc[0])
                                    entry_cost = hours * rate_value
                                    
                                    # Se for hora extra, multiplicar por 2
                                    if is_overtime:
                                        realized_cost += entry_cost * 2
                                    else:
                                        realized_cost += entry_cost
                        except Exception as e:
                            logging.warning(f"Erro ao calcular custo: {e}")
                
                # Adicionar dados migrados
                realized_hours += horas_migradas
                realized_cost += custo_migrado
                
                # Calcular percentuais
                hours_percentage = (realized_hours / total_hours * 100) if total_hours > 0 else 0
                cost_percentage = (realized_cost / total_cost * 100) if total_cost > 0 else 0
                
                # Calcular CPI (Cost Performance Index)
                # Valor Planejado = % tempo decorrido * custo total planejado
                planned_value = (time_percentage / 100) * total_cost
                cpi = planned_value / realized_cost if realized_cost > 0 else 1.0
                
                # Determinar nível de risco
                if cpi > 1.0:
                    risk_level = "Baixo"
                    risk_reason = "Abaixo do orçamento"
                elif cpi == 1.0:
                    risk_level = "Médio"
                    risk_reason = "Próximo do orçamento"
                else:
                    risk_level = "Alto"
                    risk_reason = "Acima do orçamento"

                
                # Adicionar análise adicional de risco baseada em desvios de cronograma
                # Tratamento especial para projetos do tipo "Bolsa Horas"
                if project['project_type'] == "Bolsa Horas":
                    # Para Bolsa Horas, alto consumo é bom (não é risco)
                    if hours_percentage >= 70:
                        risk_level = "Baixo"
                        risk_reason = f"Bom aproveitamento da bolsa: {hours_percentage:.1f}% utilizada"
                    elif hours_percentage >= 50:
                        risk_level = "Médio"
                        risk_reason = f"Aproveitamento médio da bolsa: {hours_percentage:.1f}% utilizada"
                    else:
                        risk_level = "Alto"
                        risk_reason = f"Baixo aproveitamento da bolsa: {hours_percentage:.1f}% utilizada"
                else:
                    # Lógica de risco padrão para outros tipos de projetos
                    schedule_variance = hours_percentage - time_percentage
                    
                    if schedule_variance > 15:  # Horas consumidas muito acima do esperado pelo tempo
                        risk_level = "Alto" if risk_level != "Alto" else risk_level
                        risk_reason += f". Consumo horas {schedule_variance:.1f}% acima"
                    elif schedule_variance < -15:  # Horas consumidas muito abaixo do esperado
                        if risk_level == "Baixo":
                            risk_reason += f". Consumo horas {abs(schedule_variance):.1f}% abaixo"
                        else:
                            risk_level = "Médio" if risk_level == "Alto" else risk_level
                            risk_reason += f". Apesar disso, o consumo de horas está {abs(schedule_variance):.1f}% abaixo do esperado"
                
                
                """
                # Adicionar análise adicional de risco baseada em desvios de cronograma
                schedule_variance = hours_percentage - time_percentage
                
                if schedule_variance > 15:  # Horas consumidas muito acima do esperado pelo tempo
                    risk_level = "Alto" if risk_level != "Alto" else risk_level
                    risk_reason += f". Consumo de horas está {schedule_variance:.1f}% acima do esperado para o cronograma"
                elif schedule_variance < -15:  # Horas consumidas muito abaixo do esperado
                    if risk_level == "Baixo":
                        risk_reason += f". Consumo de horas está {abs(schedule_variance):.1f}% abaixo do esperado para o cronograma"
                    else:
                        risk_level = "Médio" if risk_level == "Alto" else risk_level
                        risk_reason += f". Apesar disso, o consumo de horas está {abs(schedule_variance):.1f}% abaixo do esperado"
                """
                # Obter nome do cliente
                client_id = project['client_id']
                client_name = clients_df[clients_df['client_id'] == client_id]['name'].iloc[0] if not clients_df.empty else "Cliente Desconhecido"
                
                # Adicionar ao array de indicadores
                project_indicators.append({
                    'project_id': project['project_id'],
                    'project_name': project['project_name'],
                    'client_name': client_name,
                    'project_type': project['project_type'],
                    'status': project['status'],
                    'start_date': project['start_date'],
                    'end_date': project['end_date'],
                    'realized_hours': realized_hours,
                    'total_hours': total_hours,
                    'hours_percentage': hours_percentage,
                    'realized_cost': realized_cost,
                    'total_cost': total_cost,
                    'cost_percentage': cost_percentage,
                    'time_percentage': time_percentage,
                    'cpi': cpi,
                    'risk_level': risk_level,
                    'risk_reason': risk_reason,
                    'period_hours': period_hours,
                    'period_cost': period_cost
                })
            except Exception as e:
                logging.warning(f"Erro ao processar projeto {project['project_name']}: {e}")
        
        return project_indicators
    
    except Exception as e:
        import traceback
        logging.error(f"Erro ao obter indicadores de projetos: {str(e)}")
        logging.error(traceback.format_exc())
        return []# project_email_report.py

def send_email(
    recipients, 
    subject, 
    message, 
    pdf_path, 
    excel_path, 
    smtp_server, 
    smtp_port, 
    smtp_user, 
    smtp_password, 
    use_tls
):
    """
    Envia email com os relatórios anexados
    """
    try:
        # Criar email MIME
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = ", ".join(recipients)
        msg['Subject'] = subject
        
        # Adicionar corpo do email
        msg.attach(MIMEText(message, 'plain'))
        
        # Verificar se os arquivos existem antes de anexar
        has_attachments = False
        
        # Anexar PDF se existir
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as pdf_file:
                attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
                attachment.add_header(
                    'Content-Disposition', 
                    'attachment', 
                    filename=os.path.basename(pdf_path)
                )
                msg.attach(attachment)
                has_attachments = True
                st.success(f"PDF anexado ao email.")
        else:
            st.warning(f"Arquivo PDF não encontrado em {pdf_path}")
        
        # Anexar Excel se existir
        if excel_path and os.path.exists(excel_path):
            with open(excel_path, "rb") as excel_file:
                attachment = MIMEApplication(excel_file.read(), _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                attachment.add_header(
                    'Content-Disposition', 
                    'attachment', 
                    filename=os.path.basename(excel_path)
                )
                msg.attach(attachment)
                has_attachments = True
                st.success(f"Excel anexado ao email.")
        else:
            st.warning(f"Arquivo Excel não encontrado em {excel_path}")
            
        if not has_attachments:
            st.error("Nenhum anexo disponível para enviar. Verifique se os relatórios foram gerados corretamente.")
            return False
        
        # Configurar e enviar email
        server = smtplib.SMTP(smtp_server, smtp_port)
        
        if use_tls:
            server.starttls()
        
        if smtp_user and smtp_password:
            server.login(smtp_user, smtp_password)
        
        server.send_message(msg)
        server.quit()
        
        st.success(f"Email enviado com sucesso para {', '.join(recipients)}!")
        return True
    except Exception as e:
        import traceback
        error_msg = f"Erro ao enviar email: {str(e)}"
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        st.error(error_msg)
        return False


def generate_project_pdf_report(
    output_path,
    db_manager,
    annual_target_manager,
    start_date,
    end_date,
    selected_teams,
    selected_clients,
    selected_project_types,
    show_financial=True,
    show_hour_details=True
):
    """
    Gera relatório PDF com indicadores de projetos - Versão completa corrigida
    """
    logging.info("Iniciando a geração do relatório PDF de projetos.")
    
    try:
        # Inicializar PDF
        class PDF(FPDF):
            def header(self):
                # Logo apenas a partir da segunda página
                if self.page_no() != 1:  # Não mostrar o logo na primeira página (capa)
                    if os.path.exists('logo.png'):
                        self.image('logo.png', 10, 8, 33)
                    
                    # Título
                    self.set_font('Arial', 'B', 15)
                    self.cell(80)
                    self.cell(30, 25, 'Indicadores de Performance de Projetos', 0, 0, 'C')
                    self.ln(20)
            
            def footer(self):
                # Posicionar a 1.5 cm do final
                self.set_y(-15)
                # Fonte Arial itálico 8
                self.set_font('Arial', 'I', 8)
                # Número da página
                self.cell(0, 10, f'Pagina {self.page_no()}/{{nb}}', 0, 0, 'C')
        
        pdf = PDF()
        pdf.alias_nb_pages()
        
        # Capa do relatório
        pdf.add_page()

        # Adicionar o logo apenas uma vez na capa
        if os.path.exists('logo.png'):
            pdf.image('logo.png', x=10, y=20, w=60)

        pdf.set_font('Arial', 'B', 20)
        pdf.set_xy(0, 100)
        pdf.cell(210, 20, 'Relatorio - Performance de Projetos', 0, 1, 'C')
        
        # Período do relatório
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 15, f"Periodo: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}", 0, 1, 'C')
        
        # Filtros aplicados
        pdf.set_font('Arial', '', 12)
        pdf.ln(10)
        
        teams_str = ', '.join(selected_teams) if "Todas" not in selected_teams else "Todas as equipes"
        clients_str = ', '.join(selected_clients) if "Todos" not in selected_clients else "Todos os clientes"
        types_str = ', '.join(selected_project_types) if "Todos" not in selected_project_types else "Todos os tipos"
        
        pdf.set_x(40)
        pdf.cell(0, 8, f"Equipes: {teams_str}", 0, 1)
        pdf.set_x(40)
        pdf.cell(0, 8, f"Clientes: {clients_str}", 0, 1)
        pdf.set_x(40)
        pdf.cell(0, 8, f"Tipos de Projeto: {types_str}", 0, 1)
        
        # Data de geração
        pdf.set_font('Arial', 'I', 10)
        pdf.set_y(-30)
        pdf.cell(0, 10, f"Gerado automaticamente em {datetime.now().strftime('%d/%m/%Y as %H:%M')}", 0, 1, 'C')
        
        # Obter indicadores de projetos
        projetos = get_project_indicators(
            db_manager,
            annual_target_manager,
            start_date,
            end_date,
            selected_teams,
            selected_clients,
            selected_project_types
        )
        
        logging.info(f"Projetos obtidos: {len(projetos) if projetos else 0}")
        
        if not projetos:
            # Página de erro
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Sem Dados Disponíveis', 0, 1, 'C')
            
            pdf.set_font('Arial', '', 12)
            pdf.ln(10)
            pdf.multi_cell(0, 8, "Não foram encontrados dados de projetos para o período e filtros selecionados.")
            
            # Salvar o PDF mesmo assim
            pdf.output(output_path)
            return True
            
        # Converter para DataFrame
        df = pd.DataFrame(projetos)
        
        # Separar projetos do tipo "Bolsa Horas" dos demais
        bolsa_horas_df = df[df['project_type'] == 'Bolsa Horas'].copy()
        outros_projetos_df = df[df['project_type'] != 'Bolsa Horas'].copy()
        
        # Resumo executivo
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Resumo Executivo', 0, 1)
        
        # Calcular dias úteis do período
        try:
            dias_uteis = calcular_dias_uteis_projeto(start_date.date(), end_date.date())
        except:
            dias_uteis = 0
        
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 8, f"Este relatório apresenta uma visão consolidada dos indicadores de desempenho dos projetos no período de {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}.")
        
        # Dados resumidos para o sumário
        total_projetos = len(df)
        projetos_alto_risco = sum(df['risk_level'] == 'Alto')
        projetos_medio_risco = sum(df['risk_level'] == 'Médio')
        projetos_baixo_risco = sum(df['risk_level'] == 'Baixo')
        percentual_horas_media = df['hours_percentage'].mean() if 'hours_percentage' in df.columns else 0
        
        if show_financial and 'cost_percentage' in df.columns:
            percentual_custo_media = df['cost_percentage'].mean()
            cpi_medio = df['cpi'].mean()
        
        # Adicionar tabela de resumo
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, "Resumo dos Indicadores:", 0, 1)
        
        # Cabeçalho da tabela
        pdf.set_fill_color(31, 119, 180)  # Azul
        pdf.set_text_color(255, 255, 255)  # Branco
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(100, 8, "Métrica", 1, 0, 'L', True)
        pdf.cell(35, 8, "Valor", 1, 1, 'C', True)
        
        # Dados da tabela com cores alternadas
        pdf.set_text_color(0, 0, 0)  # Preto
        
        # Total de projetos
        pdf.set_font('Arial', '', 10)
        pdf.cell(100, 8, "Total de Projetos", 1, 0, 'L')
        pdf.cell(35, 8, str(total_projetos), 1, 1, 'C')
        
        # Distribuição de risco
        pdf.set_fill_color(245, 245, 245)  # Cinza claro
        pdf.cell(100, 8, "Projetos em Alto Risco", 1, 0, 'L', True)
        pdf.cell(35, 8, f"{projetos_alto_risco} ({(projetos_alto_risco/total_projetos*100):.1f}%)", 1, 1, 'C', True)
        
        pdf.cell(100, 8, "Projetos em Médio Risco", 1, 0, 'L')
        pdf.cell(35, 8, f"{projetos_medio_risco} ({(projetos_medio_risco/total_projetos*100):.1f}%)", 1, 1, 'C')
        
        pdf.set_fill_color(245, 245, 245)  # Cinza claro
        pdf.cell(100, 8, "Projetos em Baixo Risco", 1, 0, 'L', True)
        pdf.cell(35, 8, f"{projetos_baixo_risco} ({(projetos_baixo_risco/total_projetos*100):.1f}%)", 1, 1, 'C', True)
        
        # Percentual médio de horas gastas
        pdf.cell(100, 8, "Percentual Médio de Horas Gastas", 1, 0, 'L')
        pdf.cell(35, 8, f"{percentual_horas_media:.1f}%", 1, 1, 'C')
        
        # Informações financeiras (se solicitado)
        if show_financial and 'cost_percentage' in df.columns and 'cpi' in df.columns:
            pdf.set_fill_color(245, 245, 245)  # Cinza claro
            pdf.cell(100, 8, "Percentual Médio de Custo Gasto", 1, 0, 'L', True)
            pdf.cell(35, 8, f"{percentual_custo_media:.1f}%", 1, 1, 'C', True)
            
            pdf.cell(100, 8, "CPI Médio", 1, 0, 'L')
            pdf.cell(35, 8, f"{cpi_medio:.2f}", 1, 1, 'C')

        # Projetos em alto risco (apenas não-Bolsa Horas)
        high_risk_outros = outros_projetos_df[outros_projetos_df['risk_level'] == 'Alto'] if not outros_projetos_df.empty else pd.DataFrame()
        
        if not high_risk_outros.empty:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 14)
            pdf.set_fill_color(255, 230, 230)  # Vermelho claro
            pdf.cell(0, 10, "Projetos em Alto Risco", 0, 1, 'L', True)
            
            # Listar projetos em alto risco
            pdf.ln(5)
            
            high_risk_projects = high_risk_outros.sort_values('cpi')
            
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 9)
            
            # Colunas da tabela
            if show_financial:
                pdf.cell(45, 8, "Projeto", 1, 0, 'L', True)
                pdf.cell(35, 8, "Cliente", 1, 0, 'L', True)
                pdf.cell(25, 8, "% Horas", 1, 0, 'C', True)
                pdf.cell(25, 8, "% Custo", 1, 0, 'C', True)
                pdf.cell(25, 8, "CPI", 1, 0, 'C', True)
                pdf.cell(35, 8, "Razão", 1, 1, 'L', True)
            else:
                pdf.cell(100, 8, "Projeto", 1, 0, 'L', True)
                pdf.cell(90, 8, "Razão", 1, 1, 'L', True)
            
            # Dados da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 8)
            
            for idx, row in high_risk_projects.iterrows():
                # Alternar cores de fundo
                if idx % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)  # Branco
                    fill = False
                else:
                    pdf.set_fill_color(245, 245, 245)  # Cinza claro
                    fill = True
                    
                # Verificar espaço disponível
                if pdf.get_y() > 250:  # Se estiver próximo ao fim da página
                    pdf.add_page()
                    
                    # Repetir o cabeçalho
                    pdf.set_fill_color(31, 119, 180)  # Azul
                    pdf.set_text_color(255, 255, 255)  # Branco
                    pdf.set_font('Arial', 'B', 9)
                    
                    if show_financial:
                        pdf.cell(45, 8, "Projeto", 1, 0, 'L', True)
                        pdf.cell(35, 8, "Cliente", 1, 0, 'L', True)
                        pdf.cell(25, 8, "% Horas", 1, 0, 'C', True)
                        pdf.cell(25, 8, "% Custo", 1, 0, 'C', True)
                        pdf.cell(25, 8, "CPI", 1, 0, 'C', True)
                        pdf.cell(35, 8, "Razão", 1, 1, 'L', True)
                    else:
                        pdf.cell(100, 8, "Projeto", 1, 0, 'L', True)
                        pdf.cell(90, 8, "Razão", 1, 1, 'L', True)
                    
                    pdf.set_text_color(0, 0, 0)  # Preto
                    pdf.set_font('Arial', '', 8)
                
                # Truncar razão para evitar textos muito longos
                razao = str(row.get('risk_reason', 'N/A'))
                razao_truncada = (razao[:39] + '...') if len(razao) > 40 else razao
                
                if show_financial:
                    pdf.cell(45, 8, str(row['project_name'])[:20], 1, 0, 'L', fill)
                    pdf.cell(35, 8, str(row['client_name'])[:15], 1, 0, 'L', fill)
                    pdf.cell(25, 8, f"{row['hours_percentage']:.1f}%", 1, 0, 'C', fill)
                    pdf.cell(25, 8, f"{row['cost_percentage']:.1f}%", 1, 0, 'C', fill)
                    pdf.cell(25, 8, f"{row['cpi']:.2f}", 1, 0, 'C', fill)
                    pdf.cell(35, 8, razao_truncada, 1, 1, 'L', fill)
                else:
                    pdf.cell(100, 8, str(row['project_name'])[:40], 1, 0, 'L', fill)
                    pdf.cell(90, 8, razao_truncada, 1, 1, 'L', fill)

        # Tabela única com TODOS os projetos - incluindo tipo e cores de risco
        if not df.empty:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 12)  # Fonte menor para o título
            pdf.cell(0, 8, "Tabela Completa de Projetos", 0, 1)  # Altura menor para o título
            
            # Ordenar: Bolsa Horas por % de horas (desc), outros por nome do projeto
            bolsa_horas_sorted = bolsa_horas_df.sort_values('hours_percentage', ascending=False) if not bolsa_horas_df.empty else pd.DataFrame()
            outros_sorted = outros_projetos_df.sort_values(['project_name']) if not outros_projetos_df.empty else pd.DataFrame()
            
            # Concatenar: primeiro Bolsa Horas, depois outros projetos
            df_sorted = pd.concat([bolsa_horas_sorted, outros_sorted], ignore_index=True)
            
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 8)
            
            # Colunas com tipo de projeto
            col_widths = [30, 20, 15, 15, 15, 15, 15, 15, 15, 10, 10]
            pdf.cell(col_widths[0], 8, "Projeto", 1, 0, 'L', True)
            pdf.cell(col_widths[1], 8, "Cliente", 1, 0, 'L', True)
            pdf.cell(col_widths[2], 8, "Tipo", 1, 0, 'C', True)
            pdf.cell(col_widths[3], 8, "Dt.Inic", 1, 0, 'C', True)
            pdf.cell(col_widths[4], 8, "Dt.Term", 1, 0, 'C', True)
            pdf.cell(col_widths[5], 8, "H.Real", 1, 0, 'C', True)
            pdf.cell(col_widths[6], 8, "H.Plan", 1, 0, 'C', True)
            pdf.cell(col_widths[7], 8, "C.Real", 1, 0, 'C', True)
            pdf.cell(col_widths[8], 8, "C.Plan", 1, 0, 'C', True)
            pdf.cell(col_widths[9], 8, "CPI", 1, 0, 'C', True)
            pdf.cell(col_widths[10], 8, "%H", 1, 1, 'C', True)
            
            # Dados da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 7)
            
            # Contador para alternar cores
            row_count = 0
            
            for _, row in df_sorted.iterrows():
                # Verificar espaço disponível para nova página - evitar quebrar tabela
                if pdf.get_y() > 260:
                    pdf.add_page()
                    
                    # Repetir o cabeçalho
                    pdf.set_font('Arial', 'B', 12)
                    pdf.cell(0, 8, "Tabela Completa de Projetos (continuação)", 0, 1)
                    
                    pdf.set_fill_color(31, 119, 180)  # Azul
                    pdf.set_text_color(255, 255, 255)  # Branco
                    pdf.set_font('Arial', 'B', 8)
                    
                    pdf.cell(col_widths[0], 8, "Projeto", 1, 0, 'L', True)
                    pdf.cell(col_widths[1], 8, "Cliente", 1, 0, 'L', True)
                    pdf.cell(col_widths[2], 8, "Tipo", 1, 0, 'C', True)
                    pdf.cell(col_widths[3], 8, "Dt.Inic", 1, 0, 'C', True)
                    pdf.cell(col_widths[4], 8, "Dt.Term", 1, 0, 'C', True)
                    pdf.cell(col_widths[5], 8, "H.Real", 1, 0, 'C', True)
                    pdf.cell(col_widths[6], 8, "H.Plan", 1, 0, 'C', True)
                    pdf.cell(col_widths[7], 8, "C.Real", 1, 0, 'C', True)
                    pdf.cell(col_widths[8], 8, "C.Plan", 1, 0, 'C', True)
                    pdf.cell(col_widths[9], 8, "CPI", 1, 0, 'C', True)
                    pdf.cell(col_widths[10], 8, "%H", 1, 1, 'C', True)
                    
                    pdf.set_text_color(0, 0, 0)  # Preto
                    pdf.set_font('Arial', '', 7)
                
                # Alternar cores de fundo
                if row_count % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)  # Branco
                    fill = False
                else:
                    pdf.set_fill_color(245, 245, 245)  # Cinza claro
                    fill = True
                
                # Aplicar cores de risco baseadas no tipo de projeto
                project_type = str(row['project_type'])
                
                if project_type == 'Bolsa Horas':
                    # Para Bolsa Horas: lógica invertida (>=70% é bom = verde)
                    if row['hours_percentage'] >= 70:
                        pdf.set_text_color(0, 128, 0)  # Verde - Baixo risco (bom aproveitamento)
                    elif row['hours_percentage'] >= 50:
                        pdf.set_text_color(255, 193, 7)  # Amarelo - Médio risco
                    else:
                        pdf.set_text_color(255, 0, 0)  # Vermelho - Alto risco (baixo aproveitamento)
                else:
                    # Para outros tipos: lógica normal baseada no risk_level
                    if row['risk_level'] == 'Alto':
                        pdf.set_text_color(255, 0, 0)  # Vermelho - Alto risco
                    elif row['risk_level'] == 'Médio':
                        pdf.set_text_color(255, 193, 7)  # Amarelo - Médio risco
                    else:
                        pdf.set_text_color(0, 128, 0)  # Verde - Baixo risco
                
                # Truncar nomes longos
                project_name = str(row['project_name'])
                if len(project_name) > 18:
                    project_name = project_name[:17] + '...'
                    
                client_name = str(row['client_name'])
                if len(client_name) > 12:
                    client_name = client_name[:11] + '...'
                
                # Abreviar tipo de projeto
                tipo_abrev = {
                    'Bolsa Horas': 'BH',
                    'Desenvolvimento': 'Dev',
                    'Manutenção': 'Manut',
                    'Consultoria': 'Cons',
                    'OutTasking': 'OutT',
                    'Projeto': 'Proj'
                }.get(project_type, project_type[:6])
                
                # Formatar datas de forma segura
                try:
                    start_date_str = row['start_date'].strftime('%d/%m/%y') if hasattr(row['start_date'], 'strftime') else "-"
                except:
                    start_date_str = "-"
                    
                try:
                    end_date_str = row['end_date'].strftime('%d/%m/%y') if hasattr(row['end_date'], 'strftime') else "-"
                except:
                    end_date_str = "-"
                
                # Preencher os dados
                pdf.cell(col_widths[0], 7, project_name, 1, 0, 'L', fill)
                pdf.cell(col_widths[1], 7, client_name, 1, 0, 'L', fill)
                pdf.cell(col_widths[2], 7, tipo_abrev, 1, 0, 'C', fill)
                pdf.cell(col_widths[3], 7, start_date_str, 1, 0, 'C', fill)
                pdf.cell(col_widths[4], 7, end_date_str, 1, 0, 'C', fill)
                pdf.cell(col_widths[5], 7, f"{row['realized_hours']:.0f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[6], 7, f"{row['total_hours']:.0f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[7], 7, f"{row['realized_cost']:.0f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[8], 7, f"{row['total_cost']:.0f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[9], 7, f"{row['cpi']:.2f}", 1, 0, 'C', fill)
                pdf.cell(col_widths[10], 7, f"{row['hours_percentage']:.0f}%", 1, 1, 'C', fill)
                
                # Restaurar cor normal
                pdf.set_text_color(0, 0, 0)
                
                # Incrementar contador de linha
                row_count += 1
        
        # Legenda das abreviações e cores de risco - ATUALIZADA
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, "Legenda das Colunas:", 0, 1)
        
        pdf.set_font('Arial', '', 9)
        pdf.cell(0, 6, "Tipo: BH=Bolsa Horas, Dev=Desenvolvimento, Manut=Manutenção", 0, 1)
        pdf.cell(0, 6, "Cons=Consultoria, OutT=OutTasking, Proj=Projeto", 0, 1)
        pdf.cell(0, 6, "Dt.Inic=Data de Inicio, Dt.Term=Data de Termino", 0, 1)
        pdf.cell(0, 6, "H.Real=Horas Realizadas, H.Plan=Horas Planejadas", 0, 1)
        pdf.cell(0, 6, "C.Real=Custo Realizado, C.Plan=Custo Planejado", 0, 1)
        pdf.cell(0, 6, "%H=Percentual de Horas Consumidas", 0, 1)
        
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, "Legenda de Cores de Risco:", 0, 1)
        
        pdf.set_font('Arial', '', 9)
        
        # Cores para projetos normais
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 6, "Para projetos convencionais (baseado no CPI):", 0, 1)
        pdf.set_font('Arial', '', 9)
        
        pdf.set_text_color(255, 0, 0)  # Vermelho
        pdf.cell(0, 6, "Vermelho (Alto Risco): CPI < 1.0 - Projeto acima do orcamento", 0, 1)
        
        pdf.set_text_color(255, 193, 7)  # Amarelo
        pdf.cell(0, 6, "Amarelo (Medio Risco): CPI = 1.0 - Projeto proximo ao orcamento", 0, 1)
        
        pdf.set_text_color(0, 128, 0)  # Verde
        pdf.cell(0, 6, "Verde (Baixo Risco): CPI > 1.0 - Projeto abaixo do orcamento", 0, 1)
        
        pdf.set_text_color(0, 0, 0)  # Preto
        pdf.ln(3)
        
        # Cores para Bolsa Horas (lógica invertida)
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 6, "Para projetos tipo Bolsa Horas (baseado no aproveitamento):", 0, 1)
        pdf.set_font('Arial', '', 9)
        
        pdf.set_text_color(0, 128, 0)  # Verde
        pdf.cell(0, 6, "Verde (Baixo Risco): >=70% das horas - Bom aproveitamento da bolsa", 0, 1)
        
        pdf.set_text_color(255, 193, 7)  # Amarelo
        pdf.cell(0, 6, "Amarelo (Medio Risco): 50-69% das horas - Aproveitamento medio", 0, 1)
        
        pdf.set_text_color(255, 0, 0)  # Vermelho
        pdf.cell(0, 6, "Vermelho (Alto Risco): <50% das horas - Baixo aproveitamento", 0, 1)
        
        # Restaurar cor do texto
        pdf.set_text_color(0, 0, 0)
        
        # Explicação do CPI
        pdf.ln(5)
        pdf.set_font('Arial', 'I', 8)
        pdf.multi_cell(0, 5, "CPI (Cost Performance Index): Razao entre o valor planejado e o custo real. Um CPI maior que 1 indica que o projeto esta gastando menos do que o planejado (positivo), enquanto um CPI menor que 1 indica que esta gastando mais do que o planejado (negativo).")
        
        # Detalhes de horas por projeto (se solicitado)
        if show_hour_details and len(df) > 0:
            # Obter detalhes de horas por projeto
            # Criar uma tabela por projeto com informações de recursos utilizados
            
            # Detalhes de horas por projeto (se solicitado)
            if show_hour_details and len(df) > 0:
                # Obter detalhes de horas por projeto
                # Criar uma tabela por projeto com informações de recursos utilizados
                
                # Carregamos os dados necessários apenas uma vez aqui
                timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
                users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
                
                for _, project_row in df.iterrows():
                    project_id = project_row['project_id']
                    project_name = str(project_row['project_name'])
                    
                    # Obter entradas de timesheet para este projeto
                    project_timesheet = timesheet_df[timesheet_df['project_id'] == project_id]
                    
                    if not project_timesheet.empty:
                        # Verificar se precisamos de uma nova página - evitar quebrar tabela
                        if pdf.get_y() > 220:
                            pdf.add_page()
                        else:
                            pdf.ln(10)
                        
                        # Título do projeto
                        pdf.set_font('Arial', 'B', 12)
                        pdf.set_fill_color(220, 220, 220)  # Cinza claro
                        pdf.cell(180, 8, f"Detalhamento de Horas - {project_name[:50]}", 1, 1, 'L', True)  # Altura menor
                        
                        # Agrupar entradas por usuário
                        user_hours = {}
                        for _, entry in project_timesheet.iterrows():
                            user_id = entry['user_id']
                            hours = float(entry['hours'])
                            
                            if user_id not in user_hours:
                                user_hours[user_id] = {'regular': 0, 'extra': 0}
                            
                            # Verificar se é hora extra
                            is_overtime = entry.get('overtime', False)
                            if isinstance(is_overtime, (int, float)):
                                is_overtime = bool(is_overtime)
                            elif isinstance(is_overtime, str):
                                is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                            
                            if is_overtime:
                                user_hours[user_id]['extra'] += hours
                            else:
                                user_hours[user_id]['regular'] += hours
                        
                        # Total de horas do projeto
                        total_hours = sum([(u['regular'] + u['extra']) for u in user_hours.values()])
                        
                        # Criar tabela
                        pdf.set_font('Arial', 'B', 9)
                        pdf.set_fill_color(31, 119, 180)  # Azul
                        pdf.set_text_color(255, 255, 255)  # Branco
                        
                        # Cabeçalho da tabela
                        pdf.cell(60, 8, "Colaborador", 1, 0, 'L', True)
                        pdf.cell(30, 8, "Horas Normais", 1, 0, 'C', True)
                        pdf.cell(30, 8, "Horas Extras", 1, 0, 'C', True)
                        pdf.cell(30, 8, "Total", 1, 0, 'C', True)
                        pdf.cell(30, 8, "% do Projeto", 1, 1, 'C', True)
                        
                        # Linhas da tabela
                        pdf.set_font('Arial', '', 8)
                        pdf.set_text_color(0, 0, 0)  # Preto
                        
                        row_count = 0
                        for user_id, hours_data in user_hours.items():
                            # Verificar espaço para evitar quebrar tabela
                            if pdf.get_y() > 260:
                                pdf.add_page()
                                
                                # Repetir título e cabeçalho
                                pdf.set_font('Arial', 'B', 12)
                                pdf.set_fill_color(220, 220, 220)
                                pdf.cell(0, 8, f"Detalhamento de Horas - {project_name[:40]} (cont.)", 0, 1, 'L', True)
                                
                                pdf.set_font('Arial', 'B', 9)
                                pdf.set_fill_color(31, 119, 180)
                                pdf.set_text_color(255, 255, 255)
                                
                                pdf.cell(60, 8, "Colaborador", 1, 0, 'L', True)
                                pdf.cell(30, 8, "Horas Normais", 1, 0, 'C', True)
                                pdf.cell(30, 8, "Horas Extras", 1, 0, 'C', True)
                                pdf.cell(30, 8, "Total", 1, 0, 'C', True)
                                pdf.cell(30, 8, "% do Projeto", 1, 1, 'C', True)
                                
                                pdf.set_font('Arial', '', 8)
                                pdf.set_text_color(0, 0, 0)
                            
                            # Obter nome do usuário
                            user_info = users_df[users_df['user_id'] == user_id]
                            if not user_info.empty:
                                user_name = f"{user_info['First_Name'].iloc[0]} {user_info['Last_Name'].iloc[0]}"
                            else:
                                user_name = f"Usuário ID: {user_id}"
                            
                            # Calcular total e percentual
                            user_total = hours_data['regular'] + hours_data['extra']
                            percentage = (user_total / total_hours * 100) if total_hours > 0 else 0
                            
                            # Alternar cores
                            if row_count % 2 == 0:
                                pdf.set_fill_color(255, 255, 255)  # Branco
                                fill = False
                            else:
                                pdf.set_fill_color(245, 245, 245)  # Cinza claro
                                fill = True
                            
                            # Dados
                            pdf.cell(60, 7, user_name[:25], 1, 0, 'L', fill)
                            pdf.cell(30, 7, f"{hours_data['regular']:.1f}h", 1, 0, 'C', fill)
                            pdf.cell(30, 7, f"{hours_data['extra']:.1f}h", 1, 0, 'C', fill)
                            pdf.cell(30, 7, f"{user_total:.1f}h", 1, 0, 'C', fill)
                            pdf.cell(30, 7, f"{percentage:.1f}%", 1, 1, 'C', fill)
                            
                            row_count += 1
        
        # Salvar o PDF
        logging.info(f"Salvando PDF em: {output_path}")
        pdf.output(output_path)
        logging.info("PDF gerado com sucesso.")
        
        return True
    
    except Exception as e:
        import traceback
        error_msg = f"Erro ao gerar PDF: {str(e)}"
        logging.error(error_msg)
        logging.error(f"Traceback completo:\n{traceback.format_exc()}")
        return False


def generate_project_excel_report(
    output_path,
    projetos,
    start_date,
    end_date,
    selected_teams,
    selected_clients,
    selected_project_types,
    show_financial=True,
    show_hour_details=True
):
    """
    Gera relatório Excel com indicadores de projetos
    """
    try:
        # Converter lista de projetos para DataFrame
        if not projetos:
            # Criar DataFrame vazio se não houver dados
            df = pd.DataFrame()
        else:
            df = pd.DataFrame(projetos)
        
        # Criar arquivo Excel com múltiplas abas
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Aba 1: Resumo
            summary_data = {
                'Métrica': [
                    'Período do Relatório',
                    'Total de Projetos',
                    'Projetos Alto Risco',
                    'Projetos Médio Risco', 
                    'Projetos Baixo Risco',
                    'Percentual Médio de Horas Gastas',
                    'Equipes Filtradas',
                    'Clientes Filtrados',
                    'Tipos de Projeto Filtrados'
                ],
                'Valor': [
                    f"{start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}",
                    len(df) if not df.empty else 0,
                    sum(df['risk_level'] == 'Alto') if not df.empty else 0,
                    sum(df['risk_level'] == 'Médio') if not df.empty else 0,
                    sum(df['risk_level'] == 'Baixo') if not df.empty else 0,
                    f"{df['hours_percentage'].mean():.1f}%" if not df.empty else "0%",
                    ', '.join(selected_teams) if "Todas" not in selected_teams else "Todas",
                    ', '.join(selected_clients) if "Todos" not in selected_clients else "Todos",
                    ', '.join(selected_project_types) if "Todos" not in selected_project_types else "Todos"
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Resumo', index=False)
            
            # Aba 2: Dados dos Projetos
            if not df.empty:
                # Selecionar e reorganizar colunas para melhor visualização
                project_columns = [
                    'project_name', 'client_name', 'project_type', 'status',
                    'start_date', 'end_date', 'realized_hours', 'total_hours',
                    'hours_percentage', 'realized_cost', 'total_cost', 'cost_percentage',
                    'cpi', 'risk_level', 'risk_reason'
                ]
                
                # Verificar quais colunas existem no DataFrame
                available_columns = [col for col in project_columns if col in df.columns]
                
                project_data = df[available_columns].copy()
                
                # Renomear colunas para português
                column_names = {
                    'project_name': 'Nome do Projeto',
                    'client_name': 'Cliente',
                    'project_type': 'Tipo de Projeto',
                    'status': 'Status',
                    'start_date': 'Data de Início',
                    'end_date': 'Data de Término',
                    'realized_hours': 'Horas Realizadas',
                    'total_hours': 'Horas Planejadas',
                    'hours_percentage': '% Horas Consumidas',
                    'realized_cost': 'Custo Realizado',
                    'total_cost': 'Custo Planejado',
                    'cost_percentage': '% Custo Gasto',
                    'cpi': 'CPI',
                    'risk_level': 'Nível de Risco',
                    'risk_reason': 'Razão do Risco'
                }
                
                project_data = project_data.rename(columns=column_names)
                project_data.to_excel(writer, sheet_name='Dados dos Projetos', index=False)
                
                # Aba 3: Projetos por Risco
                if 'risk_level' in df.columns:
                    risk_summary = df.groupby('risk_level').agg({
                        'project_name': 'count',
                        'realized_hours': 'sum',
                        'total_hours': 'sum',
                        'realized_cost': 'sum',
                        'total_cost': 'sum'
                    }).reset_index()
                    
                    risk_summary.columns = [
                        'Nível de Risco', 'Quantidade de Projetos', 'Horas Realizadas',
                        'Horas Planejadas', 'Custo Realizado', 'Custo Planejado'
                    ]
                    
                    risk_summary.to_excel(writer, sheet_name='Projetos por Risco', index=False)
            
            else:
                # Se não houver dados, criar aba vazia com mensagem
                empty_df = pd.DataFrame({'Mensagem': ['Não há dados disponíveis para os filtros selecionados']})
                empty_df.to_excel(writer, sheet_name='Dados dos Projetos', index=False)
        
        return True
    
    except Exception as e:
        logging.error(f"Erro ao gerar Excel: {str(e)}")
        return False
