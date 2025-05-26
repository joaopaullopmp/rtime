# executive_dashboard_email.py
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import smtplib
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from datetime import datetime, timedelta
import calendar
from fpdf import FPDF
import base64
from database_manager import DatabaseManager, UserManager, ProjectManager, ClientManager, GroupManager
from annual_targets import AnnualTargetManager
from billing_manager import BillingManager
from collaborator_targets import CollaboratorTargetCalculator
from report_utils import get_feriados_portugal, calcular_dias_uteis_projeto

def executive_dashboard_email():
    """
    Módulo para geração e envio de relatórios executivos consolidados dos indicadores
    de colaboradores, projetos e faturação por email
    """
    st.title("📧 Relatório Executivo de Indicadores")
    
    # Inicializar gerenciadores
    db_manager = DatabaseManager()
    annual_target_manager = AnnualTargetManager()
    collaborator_target_calculator = CollaboratorTargetCalculator()
    billing_manager = BillingManager()
    
    # Verificar se o usuário é administrador
    if st.session_state.user_info['role'].lower() != 'admin':
        st.warning("Esta funcionalidade é exclusiva para administradores.")
        return
    
    # Carregamento de dados
    try:
        clients_df = db_manager.query_to_df("SELECT * FROM clients WHERE active = 1")
        projects_df = db_manager.query_to_df("SELECT * FROM projects")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores WHERE active = 1")
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        groups_df = db_manager.query_to_df("SELECT * FROM groups WHERE active = 1")
        rates_df = db_manager.query_to_df("SELECT * FROM rates")
        
        # Converter datas
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
            f"Relatório Executivo de Indicadores - {datetime.now().strftime('%d/%m/%Y')}"
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
            
            # Seleção de conteúdo do relatório
            st.markdown("**Seções do Relatório**")
            include_collaborator = st.checkbox("Indicadores de Colaboradores", value=True)
            include_projects = st.checkbox("Indicadores de Projetos", value=True)
            include_financial = st.checkbox("Indicadores de Faturação", value=True)
        
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
            
            # Seleção de tipos de projeto (múltipla)
            project_type_options = ["Todos"] + project_types
            selected_project_types = st.multiselect(
                "Tipos de Projeto",
                options=project_type_options,
                default=["Todos"]
            )
            
            # Seleção de formato de relatório
            report_format = st.radio(
                "Formato do Relatório",
                ["PDF", "Excel", "PDF e Excel"]
            )
        
        # Conteúdo do email
        email_message = st.text_area(
            "Mensagem do Email",
            """Prezados,

Em anexo, relatório executivo de indicadores de desempenho.

Atenciosamente,
Equipe de Gestão"""
        )
        
        # SMTP settings (collapsed by default)
        with st.expander("Configurações de SMTP"):
            smtp_server = st.text_input("Servidor SMTP", "smtp.office365.com")
            smtp_port = st.number_input("Porta SMTP", value=587, step=1)
            smtp_user = st.text_input("Usuário SMTP", "notifications@grupoerre.pt")
            smtp_password = st.text_input("Senha SMTP", type="password", value="erretech@2020")
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
                pdf_path = os.path.join(temp_dir, "relatorio_executivo_indicadores.pdf")
                pdf_result = generate_pdf_report(
                    pdf_path,
                    db_manager,
                    annual_target_manager,
                    collaborator_target_calculator,
                    billing_manager,
                    start_date,
                    end_date,
                    selected_teams,
                    selected_clients,
                    selected_project_types,
                    include_collaborator,
                    include_projects,
                    include_financial
                )
                
                # Verificar se o PDF foi gerado com sucesso
                if not pdf_result or not os.path.exists(pdf_path):
                    st.error(f"Falha ao gerar o PDF. Verifique os logs para mais detalhes.")
                    pdf_path = None

            if "Excel" in report_format:
                excel_path = os.path.join(temp_dir, "relatorio_executivo_indicadores.xlsx")
                try:
                    excel_result = generate_excel_report(
                        excel_path,
                        db_manager,
                        annual_target_manager,
                        collaborator_target_calculator,
                        billing_manager,
                        start_date,
                        end_date,
                        selected_teams,
                        selected_clients,
                        selected_project_types,
                        include_collaborator,
                        include_projects,
                        include_financial
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
                            file_name="relatorio_executivo_indicadores.pdf",
                            mime="application/pdf"
                        )
                
                if excel_path and os.path.exists(excel_path):
                    with open(excel_path, "rb") as excel_file:
                        excel_bytes = excel_file.read()
                        st.download_button(
                            label="📥 Baixar Excel",
                            data=excel_bytes,
                            file_name="relatorio_executivo_indicadores.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.error("Não foi possível gerar nenhum dos relatórios. Verifique os logs para mais detalhes.")


def generate_pdf_report(
    output_path,
    db_manager,
    annual_target_manager,
    collaborator_target_calculator,
    billing_manager,
    start_date,
    end_date,
    selected_teams,
    selected_clients,
    selected_project_types,
    include_collaborator,
    include_projects,
    include_financial
):
    """
    Gera relatório PDF com indicadores executivos
    """
    try:
        # Inicializar PDF
        class PDF(FPDF):
            def header(self):
                # Logo
                if os.path.exists('logo.png'):
                    self.image('logo.png', 10, 8, 33)
                
                # Título
                self.set_font('Arial', 'B', 15)
                self.cell(80)
                self.cell(30, 10, 'Relatorio Executivo de Indicadores', 0, 0, 'C')
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
        
        if os.path.exists('logo.png'):
            pdf.image('logo.png', x=10, y=20, w=60)
        
        pdf.set_font('Arial', 'B', 20)
        pdf.set_xy(0, 100)
        pdf.cell(210, 20, 'Relatorio Executivo de Indicadores', 0, 1, 'C')
        
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
        pdf.cell(0, 10, f"Gerado em {datetime.now().strftime('%d/%m/%Y as %H:%M')}", 0, 1, 'C')
        
        # Sumário executivo
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Sumario Executivo', 0, 1)
        
        # Separar diretório temporário para gráficos
        temp_dir = os.path.dirname(output_path)
        chart_paths = []
        
        month = start_date.month
        year = start_date.year
        
        # Calculando dias úteis do período
        dias_uteis = calcular_dias_uteis_projeto(start_date.date(), end_date.date())
        
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 8, f"Este relatorio apresenta uma visao consolidada dos principais indicadores de desempenho no periodo de {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}, cobrindo {dias_uteis} dias uteis.")
        pdf.ln(5)
        
        # Obtendo dados resumidos para o sumário
        resumo_dados = {}
        
        # Indicadores de Colaboradores
        if include_collaborator:
            # Extrair indicadores de ocupação e faturabilidade
            colaboradores = get_collaborator_indicators(
                db_manager, 
                collaborator_target_calculator, 
                month, 
                year, 
                selected_teams
            )
            
            if colaboradores:
                resumo_dados['colaboradores'] = {
                    'total': len(colaboradores),
                    'ocupacao_media': sum([c['occupation_percentage'] for c in colaboradores]) / len(colaboradores),
                    'faturabilidade_media': sum([c['billable_percentage'] for c in colaboradores]) / len(colaboradores),
                    'abaixo_meta_ocupacao': sum([1 for c in colaboradores if c['occupation_percentage'] < 87.5]),
                    'abaixo_meta_faturabilidade': sum([1 for c in colaboradores if c['billable_percentage'] < 75.0])
                }
        
        # Indicadores de Projetos
        if include_projects:
            projetos = get_project_indicators(
                db_manager,
                annual_target_manager,
                month,
                year,
                selected_teams,
                selected_clients,
                selected_project_types
            )
            
            if projetos:
                resumo_dados['projetos'] = {
                    'total': len(projetos),
                    'alto_risco': sum([1 for p in projetos if p['risk_level'] == 'Alto']),
                    'medio_risco': sum([1 for p in projetos if p['risk_level'] == 'Medio']),
                    'baixo_risco': sum([1 for p in projetos if p['risk_level'] == 'Baixo']),
                    'horas_gastas_media': sum([p['hours_percentage'] for p in projetos]) / len(projetos),
                    'custo_gasto_media': sum([p['cost_percentage'] for p in projetos]) / len(projetos)
                }
        
        # Indicadores de Faturação
        if include_financial:
            faturacao = get_revenue_indicators(
                db_manager, 
                annual_target_manager,
                billing_manager,
                month, 
                year, 
                selected_teams
            )
            
            if faturacao:
                resumo_dados['faturacao'] = {
                    'mensal_percentual': sum([f['monthly_percentage'] for f in faturacao]) / len(faturacao),
                    'trimestral_percentual': sum([f['quarterly_percentage'] for f in faturacao]) / len(faturacao),
                    'anual_percentual': sum([f['annual_percentage'] for f in faturacao]) / len(faturacao)
                }
        
        # Apresentar o sumário executivo
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, "Principais Destaques:", 0, 1)
        
        pdf.set_font('Arial', '', 11)
        
        # Adicionar pontos de destaque com base nos dados
        if 'colaboradores' in resumo_dados:
            c = resumo_dados['colaboradores']
            pdf.multi_cell(0, 8, f"* {c['total']} colaboradores avaliados com ocupacao media de {c['ocupacao_media']:.1f}% e faturabilidade media de {c['faturabilidade_media']:.1f}%.")
            
            if c['abaixo_meta_ocupacao'] > 0:
                pdf.multi_cell(0, 8, f"* {c['abaixo_meta_ocupacao']} colaboradores estao abaixo da meta de ocupacao (87.5%).")
                
            if c['abaixo_meta_faturabilidade'] > 0:
                pdf.multi_cell(0, 8, f"* {c['abaixo_meta_faturabilidade']} colaboradores estao abaixo da meta de faturabilidade (75%).")
        
        if 'projetos' in resumo_dados:
            p = resumo_dados['projetos']
            pdf.multi_cell(0, 8, f"* {p['total']} projetos analisados, com {p['alto_risco']} em alto risco, {p['medio_risco']} em medio risco e {p['baixo_risco']} em baixo risco.")
            pdf.multi_cell(0, 8, f"* O percentual medio de horas gastas e de {p['horas_gastas_media']:.1f}% e de custo gasto de {p['custo_gasto_media']:.1f}%.")
        
        if 'faturacao' in resumo_dados:
            f = resumo_dados['faturacao']
            pdf.multi_cell(0, 8, f"* A realizacao media das metas financeiras e de {f['mensal_percentual']:.1f}% para o mes, {f['trimestral_percentual']:.1f}% para o trimestre e {f['anual_percentual']:.1f}% para o ano.")
        
        # Seção de Colaboradores
        if include_collaborator and 'colaboradores' in resumo_dados:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.set_fill_color(200, 220, 255)
            pdf.cell(0, 10, 'Indicadores de Colaboradores', 0, 1, 'L', True)
            
            # Gráficos de ocupação e faturabilidade
            colaboradores = get_collaborator_indicators(
                db_manager, 
                collaborator_target_calculator, 
                month, 
                year, 
                selected_teams
            )
            
            if colaboradores:
                # Converter para DataFrame
                df = pd.DataFrame(colaboradores)
                
                # Ordenar por percentual de ocupação (decrescente)
                df_ocupacao = df.sort_values('occupation_percentage', ascending=False)
                
                # Limitar para os 10 primeiros para melhor visualização
                df_ocupacao = df_ocupacao.head(10) if len(df_ocupacao) > 10 else df_ocupacao
                
                """# Criar gráfico de ocupação
                fig_ocupacao = px.bar(
                    df_ocupacao,
                    x='name',
                    y='occupation_percentage',
                    title='Top 10 - Percentual de Ocupacao por Colaborador',
                    labels={'name': 'Colaborador', 'occupation_percentage': '% Ocupacao'},
                    color='occupation_color',
                    color_discrete_map={'green': '#4CAF50', 'red': '#F44336'},
                    text='occupation_percentage'
                )
                
                # Adicionar linha de meta
                fig_ocupacao.add_shape(
                    type="line",
                    y0=87.5,
                    y1=87.5,
                    x0=-0.5,
                    x1=len(df_ocupacao) - 0.5,
                    line=dict(color="black", width=2, dash="dash"),
                )
                
                fig_ocupacao.update_layout(
                    height=400,
                    xaxis_tickangle=-45
                )
                
                # Salvar gráfico
                ocupacao_chart_path = os.path.join(temp_dir, "ocupacao_chart.png")
                fig_ocupacao.write_image(ocupacao_chart_path, width=700, height=400)
                chart_paths.append(ocupacao_chart_path)
                
                # Adicionar gráfico ao PDF
                pdf.image(ocupacao_chart_path, x=10, y=40, w=180)
                
                # Avançar para posição após o gráfico
                pdf.set_y(160)
                
                # Ordenar por percentual de faturabilidade (decrescente)
                df_faturabilidade = df.sort_values('billable_percentage', ascending=False)
                
                # Limitar para os 10 primeiros
                df_faturabilidade = df_faturabilidade.head(10) if len(df_faturabilidade) > 10 else df_faturabilidade
                
                # Criar gráfico de faturabilidade
                fig_faturabilidade = px.bar(
                    df_faturabilidade,
                    x='name',
                    y='billable_percentage',
                    title='Top 10 - Percentual de Faturabilidade por Colaborador',
                    labels={'name': 'Colaborador', 'billable_percentage': '% Faturavel'},
                    color='billable_color',
                    color_discrete_map={'green': '#4CAF50', 'red': '#F44336'},
                    text='billable_percentage'
                )
                
                # Adicionar linha de meta
                fig_faturabilidade.add_shape(
                    type="line",
                    y0=75,
                    y1=75,
                    x0=-0.5,
                    x1=len(df_faturabilidade) - 0.5,
                    line=dict(color="black", width=2, dash="dash"),
                )
                
                fig_faturabilidade.update_layout(
                    height=400,
                    xaxis_tickangle=-45
                )
                
                # Salvar gráfico
                faturabilidade_chart_path = os.path.join(temp_dir, "faturabilidade_chart.png")
                fig_faturabilidade.write_image(faturabilidade_chart_path, width=700, height=400)
                chart_paths.append(faturabilidade_chart_path)
                
                # Próxima página para o segundo gráfico se espaço insuficiente
                if pdf.get_y() > 160:
                    pdf.add_page()
                    pdf.image(faturabilidade_chart_path, x=10, y=30, w=180)
                    pdf.set_y(150)
                else:
                    pdf.image(faturabilidade_chart_path, x=10, y=pdf.get_y(), w=180)
                    pdf.set_y(pdf.get_y() + 120)
                """
                # Adicionar resumo dos indicadores de colaboradores
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, "Resumo dos Indicadores de Colaboradores:", 0, 1)
                
                pdf.set_font('Arial', '', 10)
                pdf.multi_cell(0, 7, f"Total de Colaboradores: {len(df)}")
                pdf.multi_cell(0, 7, f"Ocupacao Media: {df['occupation_percentage'].mean():.1f}%")
                pdf.multi_cell(0, 7, f"Faturabilidade Media: {df['billable_percentage'].mean():.1f}%")
                pdf.multi_cell(0, 7, f"Colaboradores Abaixo da Meta de Ocupacao: {sum(df['occupation_percentage'] < 87.5)}")
                pdf.multi_cell(0, 7, f"Colaboradores Abaixo da Meta de Faturabilidade: {sum(df['billable_percentage'] < 75.0)}")
                
                # Legenda
                pdf.ln(5)
                pdf.set_font('Arial', 'B', 10)
                pdf.cell(0, 8, "Legenda:", 0, 1)
                
                pdf.set_font('Arial', '', 10)
                pdf.set_fill_color(76, 175, 80)  # Verde
                pdf.rect(20, pdf.get_y(), 5, 5, 'F')
                pdf.set_xy(30, pdf.get_y())
                pdf.cell(100, 5, "Acima da meta", 0, 1)
                
                pdf.set_fill_color(244, 67, 54)  # Vermelho
                pdf.rect(20, pdf.get_y(), 5, 5, 'F')
                pdf.set_xy(30, pdf.get_y())
                pdf.cell(100, 5, "Abaixo da meta", 0, 1)
                
                pdf.ln(3)
                pdf.cell(0, 5, "Meta de Ocupacao: 87.5% | Meta de Faturabilidade: 75%", 0, 1)
            
            else:
                pdf.set_font('Arial', '', 12)
                pdf.multi_cell(0, 10, "Nao foram encontrados dados de colaboradores para o periodo e filtros selecionados.")
        
        # Seção de Projetos
        if include_projects:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.set_fill_color(200, 255, 200)
            pdf.cell(0, 10, 'Indicadores de Projetos', 0, 1, 'L', True)
            
            # Gráficos e indicadores de projetos
            projetos = get_project_indicators(
                db_manager,
                annual_target_manager,
                month,
                year,
                selected_teams,
                selected_clients,
                selected_project_types
            )
            
            if projetos:
                # Converter para DataFrame
                df = pd.DataFrame(projetos)
                
                # Distribuição por status de risco
                risk_counts = df['risk_level'].value_counts().reset_index()
                risk_counts.columns = ['Status', 'Quantidade']
                
                """# Criar gráfico de pizza para distribuição de risco
                fig_risk = px.pie(
                    risk_counts, 
                    values='Quantidade', 
                    names='Status', 
                    title='Distribuicao de Projetos por Status de Risco',
                    color='Status',
                    color_discrete_map={'Alto': '#FF5252', 'Medio': '#FFC107', 'Baixo': '#4CAF50'}
                )
                
                # Gráfico de risco
                risk_chart_path = os.path.join(temp_dir, "risk_distribution.png")
                fig_risk.update_layout(height=400)
                fig_risk.write_image(risk_chart_path, width=700, height=400)
                chart_paths.append(risk_chart_path)
                
                # Adicionar gráfico ao PDF
                pdf.image(risk_chart_path, x=10, y=40, w=180)
                pdf.set_y(160)
                
                # Análise de projetos por horas gastas - top 10 mais críticos
                df_sorted = df.sort_values('hours_percentage', ascending=False)
                df_top10 = df_sorted.head(10) if len(df_sorted) > 10 else df_sorted
                
                # Criar gráfico de horas gastas
                fig_hours = px.bar(
                    df_top10,
                    y='project_name',
                    x='hours_percentage',
                    title='Top 10 - Percentual de Horas Gastas por Projeto',
                    labels={'project_name': 'Projeto', 'hours_percentage': '% de Horas Gastas'},
                    color='risk_level',
                    color_discrete_map={'Alto': '#FF5252', 'Medio': '#FFC107', 'Baixo': '#4CAF50'},
                    orientation='h',
                    text='hours_percentage'
                )
                
                fig_hours.update_layout(height=400)
                
                # Salvar gráfico
                hours_chart_path = os.path.join(temp_dir, "hours_percentage.png")
                fig_hours.write_image(hours_chart_path, width=700, height=400)
                chart_paths.append(hours_chart_path)
                
                # Próxima página para o segundo gráfico
                pdf.add_page()
                pdf.image(hours_chart_path, x=10, y=30, w=180)
                pdf.set_y(150)
                """
                
                # Adicionar resumo dos indicadores de projetos
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, "Resumo dos Indicadores de Projetos:", 0, 1)
                
                pdf.set_font('Arial', '', 10)
                pdf.multi_cell(0, 7, f"Total de Projetos: {len(df)}")
                pdf.multi_cell(0, 7, f"Distribuicao por Risco: Alto: {sum(df['risk_level'] == 'Alto')}, " + 
                              f"Medio: {sum(df['risk_level'] == 'Medio')}, " + 
                              f"Baixo: {sum(df['risk_level'] == 'Baixo')}")
                pdf.multi_cell(0, 7, f"Percentual Medio de Horas Gastas: {df['hours_percentage'].mean():.1f}%")
                pdf.multi_cell(0, 7, f"Percentual Medio de Custo Gasto: {df['cost_percentage'].mean():.1f}%")
                pdf.multi_cell(0, 7, f"CPI Medio: {df['cpi'].mean():.2f}")
                
                # Legenda
                pdf.ln(5)
                pdf.set_font('Arial', 'B', 10)
                pdf.cell(0, 8, "Legenda de Status de Risco:", 0, 1)
                
                pdf.set_font('Arial', '', 10)
                pdf.set_fill_color(255, 82, 82)  # Vermelho - Alto Risco
                pdf.rect(20, pdf.get_y(), 5, 5, 'F')
                pdf.set_xy(30, pdf.get_y())
                pdf.cell(100, 5, "Alto Risco", 0, 1)
                
                pdf.set_fill_color(255, 193, 7)  # Amarelo - Médio Risco
                pdf.rect(20, pdf.get_y(), 5, 5, 'F')
                pdf.set_xy(30, pdf.get_y())
                pdf.cell(100, 5, "Medio Risco", 0, 1)
                
                pdf.set_fill_color(76, 175, 80)  # Verde - Baixo Risco
                pdf.rect(20, pdf.get_y(), 5, 5, 'F')
                pdf.set_xy(30, pdf.get_y())
                pdf.cell(100, 5, "Baixo Risco", 0, 1)
            
            else:
                pdf.set_font('Arial', '', 12)
                pdf.multi_cell(0, 10, "Nao foram encontrados dados de projetos para o periodo e filtros selecionados.")
        
        # Seção de Faturação
        if include_financial:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.set_fill_color(255, 220, 200)
            pdf.cell(0, 10, 'Indicadores de Faturacao', 0, 1, 'L', True)
            
            # Gráficos e indicadores de faturação
            faturacao = get_revenue_indicators(
                db_manager, 
                annual_target_manager,
                billing_manager,
                month, 
                year, 
                selected_teams
            )
            
            if faturacao:
                # Converter para DataFrame
                df = pd.DataFrame(faturacao)
                
            """    # Criar gráfico de barras para metas mensais
                fig_monthly = px.bar(
                    df,
                    x='company_name',
                    y='monthly_percentage',
                    title='Percentual da Meta Mensal por Empresa',
                    labels={'company_name': 'Empresa', 'monthly_percentage': '% da Meta Mensal'},
                    color='monthly_color',
                    color_discrete_map={'green': '#4CAF50', 'yellow': '#FFC107', 'red': '#F44336'},
                    text='monthly_percentage'
                )
                
                # Adicionar linha de referência em 60% e 80%
                fig_monthly.add_shape(
                    type="line",
                    y0=60,
                    y1=60,
                    x0=-0.5,
                    x1=len(df) - 0.5,
                    line=dict(color="black", width=1, dash="dot"),
                )
                
                fig_monthly.add_shape(
                    type="line",
                    y0=80,
                    y1=80,
                    x0=-0.5,
                    x1=len(df) - 0.5,
                    line=dict(color="black", width=1, dash="dot"),
                )
                
                fig_monthly.update_layout(height=400)
                
                # Salvar gráfico
                monthly_chart_path = os.path.join(temp_dir, "monthly_revenue.png")
                fig_monthly.write_image(monthly_chart_path, width=700, height=400)
                chart_paths.append(monthly_chart_path)
                
                # Adicionar gráfico ao PDF
                pdf.image(monthly_chart_path, x=10, y=40, w=180)
                pdf.set_y(160)
                
                # Criar gráfico de barras para metas anuais
                fig_annual = px.bar(
                    df,
                    x='company_name',
                    y='annual_percentage',
                    title='Percentual da Meta Anual por Empresa',
                    labels={'company_name': 'Empresa', 'annual_percentage': '% da Meta Anual'},
                    color='annual_color',
                    color_discrete_map={'green': '#4CAF50', 'yellow': '#FFC107', 'red': '#F44336'},
                    text='annual_percentage'
                )
                
                # Adicionar linha de referência em 60% e 80%
                fig_annual.add_shape(
                    type="line",
                    y0=60,
                    y1=60,
                    x0=-0.5,
                    x1=len(df) - 0.5,
                    line=dict(color="black", width=1, dash="dot"),
                )
                
                fig_annual.add_shape(
                    type="line",
                    y0=80,
                    y1=80,
                    x0=-0.5,
                    x1=len(df) - 0.5,
                    line=dict(color="black", width=1, dash="dot"),
                )
                
                fig_annual.update_layout(height=400)
                
                # Salvar gráfico
                annual_chart_path = os.path.join(temp_dir, "annual_revenue.png")
                fig_annual.write_image(annual_chart_path, width=700, height=400)
                chart_paths.append(annual_chart_path)
                
                # Próxima página para o segundo gráfico
                pdf.add_page()
                pdf.image(annual_chart_path, x=10, y=30, w=180)
                pdf.set_y(150)
                """
                
                # Adicionar resumo dos indicadores de faturação
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, "Resumo dos Indicadores de Faturacao:", 0, 1)
            
            pdf.set_font('Arial', '', 10)
            pdf.multi_cell(0, 7, f"Total de Empresas: {len(df)}")
            pdf.multi_cell(0, 7, f"Percentual Medio da Meta Mensal: {df['monthly_percentage'].mean():.1f}%")
            pdf.multi_cell(0, 7, f"Percentual Medio da Meta Trimestral: {df['quarterly_percentage'].mean():.1f}%")
            pdf.multi_cell(0, 7, f"Percentual Medio da Meta Anual: {df['annual_percentage'].mean():.1f}%")
            
            # Legenda
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Legenda:", 0, 1)
            
            pdf.set_font('Arial', '', 10)
            pdf.set_fill_color(76, 175, 80)  # Verde - ≥ 80%
            pdf.rect(20, pdf.get_y(), 5, 5, 'F')
            pdf.set_xy(30, pdf.get_y())
            pdf.cell(100, 5, ">= 80% da meta", 0, 1)
            
            pdf.set_fill_color(255, 193, 7)  # Amarelo - Entre 60% e 80%
            pdf.rect(20, pdf.get_y(), 5, 5, 'F')
            pdf.set_xy(30, pdf.get_y())
            pdf.cell(100, 5, "Entre 60% e 80% da meta", 0, 1)
            
            pdf.set_fill_color(244, 67, 54)  # Vermelho - < 60%
            pdf.rect(20, pdf.get_y(), 5, 5, 'F')
            pdf.set_xy(30, pdf.get_y())
            pdf.cell(100, 5, "< 60% da meta", 0, 1)
        
        else:
            pdf.set_font('Arial', '', 12)
            pdf.multi_cell(0, 10, "Nao foram encontrados dados de faturacao para o periodo e filtros selecionados.")
        
        # Salvar o PDF
        pdf.output(output_path)
        
        # Limpar arquivos temporários de gráficos
        for chart_path in chart_paths:
            if os.path.exists(chart_path):
                os.remove(chart_path)
        
        return True
    
    except Exception as e:
        import traceback
        error_msg = f"Erro ao gerar PDF: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return False


def generate_excel_report(
    output_path,
    db_manager,
    annual_target_manager,
    collaborator_target_calculator,
    billing_manager,
    start_date,
    end_date,
    selected_teams,
    selected_clients,
    selected_project_types,
    include_collaborator,
    include_projects,
    include_financial
):
    """
    Gera relatório Excel com indicadores executivos
    """
    try:
        # Configurações do mês e ano para as consultas
        month = start_date.month
        year = start_date.year
        
        # Criar escritor Excel
        writer = pd.ExcelWriter(output_path, engine='openpyxl')
        
        # Planilha de Resumo
        resumo_data = {
            'Período': [f"{start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}"],
            'Dias Úteis': [calcular_dias_uteis_projeto(start_date.date(), end_date.date())],
            'Data de Geração': [datetime.now().strftime('%d/%m/%Y %H:%M:%S')]
        }
        
        resumo_df = pd.DataFrame(resumo_data)
        resumo_df.to_excel(writer, sheet_name='Resumo', index=False, startrow=0)
        
        # Adicionar filtros aplicados
        filtros_data = {
            'Filtro': ['Equipes', 'Clientes', 'Tipos de Projeto'],
            'Valor': [
                ', '.join(selected_teams) if "Todas" not in selected_teams else "Todas as equipes",
                ', '.join(selected_clients) if "Todos" not in selected_clients else "Todos os clientes",
                ', '.join(selected_project_types) if "Todos" not in selected_project_types else "Todos os tipos"
            ]
        }
        
        filtros_df = pd.DataFrame(filtros_data)
        filtros_df.to_excel(writer, sheet_name='Resumo', index=False, startrow=4)
        
        # Planilha de Colaboradores
        if include_collaborator:
            colaboradores = get_collaborator_indicators(
                db_manager, 
                collaborator_target_calculator, 
                month, 
                year, 
                selected_teams
            )
            
            if colaboradores:
                df = pd.DataFrame(colaboradores)
                
                # Selecionar e renomear colunas
                colabs_df = df[['name', 'occupation_percentage', 'billable_percentage', 'total_hours', 'billable_hours']]
                colabs_df.columns = ['Colaborador', 'Ocupação (%)', 'Faturabilidade (%)', 'Horas Totais', 'Horas Faturáveis']
                
                # Adicionar avaliação de meta
                colabs_df['Atinge Meta Ocupação'] = colabs_df['Ocupação (%)'] >= 87.5
                colabs_df['Atinge Meta Faturabilidade'] = colabs_df['Faturabilidade (%)'] >= 75.0
                
                # Escrever planilha
                colabs_df.to_excel(writer, sheet_name='Colaboradores', index=False)
                
                # Adicionar resumo
                resumo_colabs = {
                    'Métrica': [
                        'Total de Colaboradores',
                        'Ocupação Média (%)',
                        'Faturabilidade Média (%)',
                        'Colaboradores Abaixo da Meta de Ocupação',
                        'Colaboradores Abaixo da Meta de Faturabilidade'
                    ],
                    'Valor': [
                        len(df),
                        df['occupation_percentage'].mean(),
                        df['billable_percentage'].mean(),
                        sum(df['occupation_percentage'] < 87.5),
                        sum(df['billable_percentage'] < 75.0)
                    ]
                }
                
                resumo_colabs_df = pd.DataFrame(resumo_colabs)
                resumo_colabs_df.to_excel(writer, sheet_name='Resumo Colaboradores', index=False)
        
        # Planilha de Projetos
        if include_projects:
            projetos = get_project_indicators(
                db_manager,
                annual_target_manager,
                month,
                year,
                selected_teams,
                selected_clients,
                selected_project_types
            )
            
            if projetos:
                df = pd.DataFrame(projetos)
                
                # Selecionar e renomear colunas
                projs_df = df[[
                    'project_name', 'client_name', 'project_type', 'status',
                    'realized_hours', 'total_hours', 'hours_percentage',
                    'realized_cost', 'total_cost', 'cost_percentage',
                    'cpi', 'risk_level', 'risk_reason'
                ]]
                
                projs_df.columns = [
                    'Projeto', 'Cliente', 'Tipo', 'Status',
                    'Horas Atuais', 'Horas Planejadas', '% Horas Gastas',
                    'Custo Atual', 'Custo Planejado', '% Custo Gasto',
                    'CPI', 'Nível de Risco', 'Razão do Risco'
                ]
                
                # Escrever planilha
                projs_df.to_excel(writer, sheet_name='Projetos', index=False)
                
                # Adicionar resumo
                resumo_projs = {
                    'Métrica': [
                        'Total de Projetos',
                        'Projetos em Alto Risco',
                        'Projetos em Médio Risco',
                        'Projetos em Baixo Risco',
                        'Percentual Médio de Horas Gastas',
                        'Percentual Médio de Custo Gasto',
                        'CPI Médio'
                    ],
                    'Valor': [
                        len(df),
                        sum(df['risk_level'] == 'Alto'),
                        sum(df['risk_level'] == 'Médio'),
                        sum(df['risk_level'] == 'Baixo'),
                        df['hours_percentage'].mean(),
                        df['cost_percentage'].mean(),
                        df['cpi'].mean()
                    ]
                }
                
                resumo_projs_df = pd.DataFrame(resumo_projs)
                resumo_projs_df.to_excel(writer, sheet_name='Resumo Projetos', index=False)
                
                # Distribuição por tipo de projeto
                tipos_count = df['project_type'].value_counts().reset_index()
                tipos_count.columns = ['Tipo de Projeto', 'Quantidade']
                tipos_count.to_excel(writer, sheet_name='Tipos Projeto', index=False)
                
                # Distribuição por risco
                risk_count = df['risk_level'].value_counts().reset_index()
                risk_count.columns = ['Nível de Risco', 'Quantidade']
                risk_count.to_excel(writer, sheet_name='Níveis Risco', index=False)
        
        # Planilha de Faturação
        if include_financial:
            faturacao = get_revenue_indicators(
                db_manager, 
                annual_target_manager,
                billing_manager,
                month, 
                year, 
                selected_teams
            )
            
            if faturacao:
                df = pd.DataFrame(faturacao)
                
                # Selecionar e renomear colunas
                fat_df = df[[
                    'company_name', 'monthly_target', 'monthly_revenue', 'monthly_percentage',
                    'quarterly_target', 'quarterly_revenue', 'quarterly_percentage',
                    'annual_target', 'annual_revenue', 'annual_percentage'
                ]]
                
                fat_df.columns = [
                    'Empresa', 'Meta Mensal', 'Faturação Mensal', '% Mensal',
                    'Meta Trimestral', 'Faturação Trimestral', '% Trimestral',
                    'Meta Anual', 'Faturação Anual', '% Anual'
                ]
                
                # Escrever planilha
                fat_df.to_excel(writer, sheet_name='Faturação', index=False)
                
                # Adicionar resumo
                resumo_fat = {
                    'Métrica': [
                        'Total de Empresas',
                        'Percentual Médio da Meta Mensal',
                        'Percentual Médio da Meta Trimestral',
                        'Percentual Médio da Meta Anual',
                        'Empresas Abaixo de 60% da Meta Mensal',
                        'Empresas Abaixo de 60% da Meta Anual'
                    ],
                    'Valor': [
                        len(df),
                        df['monthly_percentage'].mean(),
                        df['quarterly_percentage'].mean(),
                        df['annual_percentage'].mean(),
                        sum(df['monthly_percentage'] < 60),
                        sum(df['annual_percentage'] < 60)
                    ]
                }
                
                resumo_fat_df = pd.DataFrame(resumo_fat)
                resumo_fat_df.to_excel(writer, sheet_name='Resumo Faturação', index=False)
        
        # Salvar o arquivo Excel
        writer._save()
        return True
    
    except Exception as e:
        import traceback
        error_msg = f"Erro ao gerar Excel: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return False


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
                    filename="relatorio_executivo_indicadores.pdf"
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
                    filename="relatorio_executivo_indicadores.xlsx"
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
        print(error_msg)
        print(traceback.format_exc())
        st.error(error_msg)
        return False


def get_collaborator_indicators(db_manager, collaborator_target_calculator, month, year, selected_teams):
    """
    Obter indicadores de colaboradores para o relatório
    """
    try:
        # Carregar dados necessários
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores WHERE active = 1")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        
        # Filtrar por equipe
        if "Todas" not in selected_teams:
            # Encontrar os usuários que pertencem às equipes selecionadas
            filtered_users = []
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
                    
                    # Verificar se o usuário pertence a alguma das equipes selecionadas
                    if any(team in user_groups for team in selected_teams):
                        filtered_users.append(user['user_id'])
                except Exception as e:
                    print(f"Erro ao processar grupos do usuário {user['First_Name']} {user['Last_Name']}: {str(e)}")
            
            # Filtrar usuários
            users_df = users_df[users_df['user_id'].isin(filtered_users)]
        
        if users_df.empty:
            return []
        
        # Calcular dias úteis do mês
        primeiro_dia = datetime(year, month, 1)
        ultimo_dia = datetime(year, month, calendar.monthrange(year, month)[1], 23, 59, 59)
        
        # Obter lista de feriados para o ano
        feriados = get_feriados_portugal(year)
        
        # Calcular dias úteis (de trabalho) no mês, excluindo feriados
        dias_uteis = 0
        data_atual = primeiro_dia
        while data_atual <= ultimo_dia:
            # Se não for sábado (5) nem domingo (6) e não for feriado
            if data_atual.weekday() < 5 and data_atual.date() not in feriados:
                dias_uteis += 1
            data_atual += timedelta(days=1)
        
        # Horas úteis totais (considerando 8 horas por dia útil)
        horas_uteis_mes = dias_uteis * 8
        
        # Para cada colaborador, calcular indicadores
        collaborator_indicators = []
        
        for _, user in users_df.iterrows():
            # Obter dados de metas do colaborador
            user_targets = collaborator_target_calculator.get_user_targets(user['user_id'], year)
            
            # Filtrar entradas de timesheet para o colaborador e mês específico
            try:
                user_timesheet = timesheet_df[
                    (timesheet_df['user_id'] == user['user_id']) & 
                    (pd.to_datetime(timesheet_df['start_date'], format='mixed').dt.month == month) &
                    (pd.to_datetime(timesheet_df['start_date'], format='mixed').dt.year == year)
                ]
            except Exception as e:
                # Fallback em caso de erro de conversão de data
                print(f"Erro ao filtrar timesheet para {user['First_Name']} {user['Last_Name']}: {e}")
                user_timesheet = pd.DataFrame()
            
            # Calcular horas realizadas (total de horas registradas)
            total_hours = user_timesheet['hours'].sum() if not user_timesheet.empty else 0
            
            # Calcular horas faturáveis (apenas entradas marcadas como billable=True)
            billable_hours = user_timesheet[user_timesheet['billable'] == True]['hours'].sum() if not user_timesheet.empty else 0
            
            # Calcular percentuais de ocupação e faturabilidade baseados nas horas realizadas
            # Ocupação = (Horas realizadas / Horas úteis do mês) * 100
            occupation_percentage = (total_hours / horas_uteis_mes * 100) if horas_uteis_mes > 0 else 0
            
            # Faturabilidade = (Horas faturáveis / Horas úteis do mês) * 100
            billable_percentage = (billable_hours / horas_uteis_mes * 100) if horas_uteis_mes > 0 else 0
            
            # Metas definidas
            target_occupation = 87.5  # Meta de ocupação (87,5%)
            target_billable = 75.0    # Meta de horas faturáveis (75.0%)
            
            # Determinar cores dos indicadores - simplificado para verde/vermelho
            occupation_color = "green" if occupation_percentage >= target_occupation else "red"
            billable_color = "green" if billable_percentage >= target_billable else "red"
            
            # Adicionar ao array de indicadores
            collaborator_indicators.append({
                "user_id": user['user_id'],
                "name": f"{user['First_Name']} {user['Last_Name']}",
                "occupation_percentage": occupation_percentage,
                "billable_percentage": billable_percentage,
                "occupation_color": occupation_color,
                "billable_color": billable_color,
                "target_occupation": target_occupation,
                "target_billable": target_billable,
                "total_hours": total_hours,  # Quantidade de horas realizadas
                "billable_hours": billable_hours  # Quantidade de horas faturáveis
            })
        
        return collaborator_indicators
    
    except Exception as e:
        import traceback
        print(f"Erro ao obter indicadores de colaboradores: {str(e)}")
        print(traceback.format_exc())
        return []
    
def get_project_indicators(db_manager, annual_target_manager, month, year, selected_teams, selected_clients, selected_project_types):
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
        
        # Filtrar por status (apenas ativos)
        filtered_projects = filtered_projects[filtered_projects["status"].str.lower() == "active"]
        
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
        
        # Definir o início e fim do mês de referência
        inicio_mes = datetime(year, month, 1)
        ultimo_dia = calendar.monthrange(year, month)[1]
        fim_mes = datetime(year, month, ultimo_dia, 23, 59, 59)
        
        # Converter datas em timesheet para datetime para facilitar filtros
        if 'start_date' in timesheet_df.columns:
            try:
                timesheet_df['start_date_dt'] = pd.to_datetime(timesheet_df['start_date'], format='mixed')
            except Exception as e:
                print(f"Erro ao converter datas: {e}")
                timesheet_df['start_date_dt'] = pd.to_datetime('2000-01-01')
        
        for _, project in filtered_projects.iterrows():
            try:
                # Obter registros de timesheet para o projeto
                project_entries = timesheet_df[timesheet_df["project_id"] == project["project_id"]]
                
                # Verificar se existem dados migrados
                horas_migradas = 0
                if 'horas_realizadas_mig' in project and not pd.isna(project['horas_realizadas_mig']):
                    horas_migradas = float(project['horas_realizadas_mig'])
                
                custo_migrado = 0
                if 'custo_realizado_mig' in project and not pd.isna(project['custo_realizado_mig']):
                    custo_migrado = float(project['custo_realizado_mig'])
                
                # Calcular dias úteis totais do projeto
                dias_uteis_projeto = calcular_dias_uteis(project['start_date'], project['end_date'])
                
                # Entradas do projeto no mês atual
                try:
                    month_entries = project_entries[
                        (pd.to_datetime(project_entries["start_date"], format='mixed') >= inicio_mes) &
                        (pd.to_datetime(project_entries["start_date"], format='mixed') <= fim_mes)
                    ]
                except Exception as e:
                    print(f"Erro ao filtrar entradas do mês: {e}")
                    month_entries = pd.DataFrame()
                
                # Calcular horas regulares e extras
                month_hours_regular = 0
                month_hours_extra = 0
                
                if not month_entries.empty:
                    month_hours_regular = month_entries[~month_entries['overtime'].astype(bool)]['hours'].sum()
                    # Para horas extras, calcular o dobro para contabilização
                    month_hours_extra_original = month_entries[month_entries['overtime'].astype(bool)]['hours'].sum()
                    month_hours_extra = month_hours_extra_original * 2
                
                # Total de horas no mês (regulares + extras*2)
                month_hours = float(month_hours_regular + month_hours_extra)
                
                # Calcular custo realizado para o mês
                month_cost = 0
                if not month_entries.empty:
                    for _, entry in month_entries.iterrows():
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
                                    month_cost += entry_cost * 2  # Dobro para horas extras
                                else:
                                    month_cost += entry_cost  # Normal para horas regulares
                        except Exception as e:
                            print(f"Erro ao processar entrada: {e}")
                
                # Calcular percentual de tempo decorrido do projeto
                total_days = (project['end_date'] - project['start_date']).days
                elapsed_days = (datetime.now() - project['start_date']).days
                elapsed_days = max(0, min(elapsed_days, total_days))  # Garantir entre 0 e total_days
                time_percentage = (elapsed_days / total_days * 100) if total_days > 0 else 0
                
                # Obter horas e custo totais planejados
                total_hours = float(project['total_hours']) if pd.notna(project['total_hours']) else 0
                total_cost = float(project['total_cost']) if pd.notna(project['total_cost']) else 0
                
                # Calcular horas e custo realizados (incluindo dados migrados)
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
                            print(f"Erro ao calcular custo: {e}")
                
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
                if cpi >= 1.1:
                    risk_level = "Baixo"
                    risk_reason = "O projeto está abaixo do orçamento para o período atual"
                elif cpi >= 0.9:
                    risk_level = "Médio"
                    risk_reason = "O projeto está próximo do orçamento projetado para o período atual"
                else:
                    risk_level = "Alto"
                    risk_reason = "O projeto está acima do orçamento para o período atual"
                
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
                    'risk_reason': risk_reason
                })
            except Exception as e:
                print(f"Erro ao processar projeto {project['project_name']}: {e}")
        
        return project_indicators
    
    except Exception as e:
        import traceback
        print(f"Erro ao obter indicadores de projetos: {str(e)}")
        print(traceback.format_exc())
        return []


def get_revenue_indicators(db_manager, annual_target_manager, billing_manager, month, year, selected_teams):
    """
    Obter indicadores de faturação para o relatório
    """
    try:
        # Carregar dados necessários
        annual_targets = annual_target_manager.read()
        invoices_df = billing_manager.get_invoice()  # Buscar todas as faturas
        clients_df = db_manager.query_to_df("SELECT * FROM clients")
        projects_df = db_manager.query_to_df("SELECT * FROM projects")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        
        # Filtrar por equipe, se necessário
        if "Todas" not in selected_teams:
            annual_targets = annual_targets[annual_targets["company_name"].isin(selected_teams)]
        
        if annual_targets.empty:
            return []
        
        # Calcular as metas para o ano atual
        year_targets = annual_targets[annual_targets["target_year"] == year]
        
        if year_targets.empty:
            return []
        
        # Definir períodos
        primeiro_dia_mes = datetime(year, month, 1)
        ultimo_dia_mes = datetime(year, month, calendar.monthrange(year, month)[1], 23, 59, 59)
        
        inicio_ano = datetime(year, 1, 1)
        fim_ano = datetime(year, 12, 31, 23, 59, 59)
        
        # Determinar trimestre atual
        current_quarter = (month - 1) // 3 + 1
        inicio_trimestre = datetime(year, ((current_quarter - 1) * 3) + 1, 1)
        fim_trimestre_mes = min(current_quarter * 3, 12)
        ultimo_dia_trimestre = calendar.monthrange(year, fim_trimestre_mes)[1]
        fim_trimestre = datetime(year, fim_trimestre_mes, ultimo_dia_trimestre, 23, 59, 59)
        
        # Calcular indicadores de faturação
        revenue_indicators = []
        
        for _, target in year_targets.iterrows():
            company_name = target["company_name"]
            annual_target = target["target_value"]
            
            # Calcular metas mensais e trimestrais
            monthly_target = annual_target / 12
            quarterly_target = annual_target / 4
            
            # Filtrar faturas por empresa usando projetos associados à empresa
            company_invoices = invoices_df.copy() if not invoices_df.empty else pd.DataFrame()
            
            if company_name != "Todas" and not company_invoices.empty:
                # Mapear projetos da empresa
                company_projects = []
                
                # Identificar o group_id correspondente à empresa
                group_id = None
                for _, group in groups_df.iterrows():
                    if group['group_name'] == company_name:
                        group_id = group['id']
                        break
                
                if group_id is not None:
                    # Filtrar projetos da empresa pelo group_id
                    company_projects_df = projects_df[projects_df['group_id'] == group_id]
                    company_projects = company_projects_df['project_id'].tolist()
                    
                    # Filtrar faturas pelos projetos da empresa
                    if company_projects:
                        company_invoices = company_invoices[company_invoices['project_id'].isin(company_projects)]
                    else:
                        company_invoices = pd.DataFrame()  # Sem projetos, sem faturas
            
            # Calcular faturação efetiva por período usando datas de pagamento das faturas
            if not company_invoices.empty:
                # Converter datas para datetime
                company_invoices['payment_date'] = pd.to_datetime(company_invoices['payment_date'])
                
                # Filtrar para o mês atual
                month_invoices = company_invoices[
                    (company_invoices['payment_date'] >= primeiro_dia_mes) & 
                    (company_invoices['payment_date'] <= ultimo_dia_mes)
                ]
                monthly_revenue = month_invoices['amount'].sum() if not month_invoices.empty else 0
                
                # Filtrar para o trimestre atual
                quarter_invoices = company_invoices[
                    (company_invoices['payment_date'] >= inicio_trimestre) & 
                    (company_invoices['payment_date'] <= fim_trimestre)
                ]
                quarterly_revenue = quarter_invoices['amount'].sum() if not quarter_invoices.empty else 0
                
                # Filtrar para o ano atual
                year_invoices = company_invoices[
                    (company_invoices['payment_date'] >= inicio_ano) & 
                    (company_invoices['payment_date'] <= fim_ano)
                ]
                annual_revenue = year_invoices['amount'].sum() if not year_invoices.empty else 0
            else:
                # Sem faturas registradas
                monthly_revenue = 0
                quarterly_revenue = 0
                annual_revenue = 0
            
            # Calcular percentuais de conclusão
            monthly_percentage = (monthly_revenue / monthly_target * 100) if monthly_target > 0 else 0
            quarterly_percentage = (quarterly_revenue / quarterly_target * 100) if quarterly_target > 0 else 0
            annual_percentage = (annual_revenue / annual_target * 100) if annual_target > 0 else 0
            
            # Determinar cores dos indicadores
            # Mensal
            if monthly_percentage >= 80:
                monthly_color = "green"
            elif monthly_percentage >= 60:
                monthly_color = "yellow"
            else:
                monthly_color = "red"
            
            # Trimestral
            if quarterly_percentage >= 80:
                quarterly_color = "green"
            elif quarterly_percentage >= 60:
                quarterly_color = "yellow"
            else:
                quarterly_color = "red"
                
            # Anual
            if annual_percentage >= 80:
                annual_color = "green"
            elif annual_percentage >= 60:
                annual_color = "yellow"
            else:
                annual_color = "red"
            
            # Adicionar ao array de indicadores
            revenue_indicators.append({
                "company_name": company_name,
                "monthly_target": monthly_target,
                "monthly_revenue": monthly_revenue,
                "monthly_percentage": monthly_percentage,
                "quarterly_target": quarterly_target,
                "quarterly_revenue": quarterly_revenue,
                "quarterly_percentage": quarterly_percentage,
                "annual_target": annual_target,
                "annual_revenue": annual_revenue,
                "annual_percentage": annual_percentage,
                "monthly_color": monthly_color,
                "quarterly_color": quarterly_color,
                "annual_color": annual_color
            })
        
        return revenue_indicators
    
    except Exception as e:
        import traceback
        print(f"Erro ao obter indicadores de faturação: {str(e)}")
        print(traceback.format_exc())
        return []