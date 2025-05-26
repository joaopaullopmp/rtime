import pandas as pd
import sqlite3
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import calendar
import streamlit as st
import io
import os
import smtplib
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from database_manager import DatabaseManager
import numpy as np
from fpdf import FPDF
import matplotlib.pyplot as plt

def format_hours_minutes(hours):
    """Converte um valor decimal de horas para o formato HH:MM"""
    if pd.isna(hours) or not isinstance(hours, (int, float)):
        return "00:00"
    hours_int = int(hours)  # Parte inteira (horas)
    minutes = int((hours - hours_int) * 60)  # Parte decimal convertida para minutos
    return f"{hours_int:02d}:{minutes:02d}"  # Formato HH:MM com zero à esquerda

class CommercialReportPDF(FPDF):
    """Classe personalizada para criar relatórios de indicadores comerciais em PDF"""
    def __init__(self, title="Relatório de Indicadores Comerciais", **kwargs):
        super().__init__(**kwargs)
        self.title = title
        self.set_auto_page_break(auto=True, margin=15)
        self.alias_nb_pages()  # Para usar {nb} no rodapé
        
    def header(self):
        # Logo
        if os.path.exists('logo.png'):
            self.image('logo.png', 10, 8, 33)
        
        # Título
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, self.title, 0, 1, 'C')
        self.ln(5)
        
    def footer(self):
        # Posição a 1.5 cm do final
        self.set_y(-15)
        # Arial itálico 8
        self.set_font('Arial', 'I', 8)
        # Número da página
        self.cell(0, 10, f'Página {self.page_no()}/{{nb}}', 0, 0, 'C')
        
    def add_heading(self, text, size=14, spacing=5):
        self.set_font('Arial', 'B', size)
        self.cell(0, 10, text, 0, 1)
        self.ln(spacing)
        
    def add_subheading(self, text, size=12, spacing=3):
        self.set_font('Arial', 'B', size)
        self.cell(0, 8, text, 0, 1)
        self.ln(spacing)
        
    def add_paragraph(self, text, size=10, spacing=5):
        self.set_font('Arial', '', size)
        self.multi_cell(0, 6, text)
        self.ln(spacing)
        
    def add_metric(self, label, value, size=10, spacing=2):
        self.set_font('Arial', 'B', size)
        self.cell(60, 8, label, 0, 0)
        self.set_font('Arial', '', size)
        self.cell(0, 8, str(value), 0, 1)
        self.ln(spacing)
        
    def add_table_header(self, headers, widths, height=7, fill=True):
        self.set_font('Arial', 'B', 9)
        self.set_fill_color(41, 128, 185)  # Azul
        self.set_text_color(255, 255, 255)  # Branco
        
        for i in range(len(headers)):
            self.cell(widths[i], height, headers[i], 1, 0, 'C', fill)
        self.ln()
        
        # Resetar cor do texto
        self.set_text_color(0, 0, 0)
        
    def add_table_row(self, data, widths, height=6, fill=False, alignments=None):
        if alignments is None:
            alignments = ['L'] * len(data)
            
        self.set_font('Arial', '', 8)
        
        if fill:
            self.set_fill_color(240, 240, 240)  # Cinza claro
            
        for i in range(len(data)):
            self.cell(widths[i], height, str(data[i]), 1, 0, alignments[i], fill)
        self.ln()
        
    def add_table(self, df, max_rows=None, title=None):
        # Adicionar título se fornecido
        if title:
            self.add_subheading(title)
            
        # Verificar se o DataFrame está vazio
        if df.empty:
            self.add_paragraph("Não há dados disponíveis.")
            return
            
        # Limitar o número de linhas se necessário
        if max_rows is not None:
            df_to_display = df.head(max_rows)
        else:
            df_to_display = df
            
        # Calcular larguras das colunas 
        n_columns = len(df_to_display.columns)
        page_width = self.w - 2*self.l_margin
        col_widths = [page_width / n_columns] * n_columns
        
        # Ajustar largura para colunas específicas se necessário (como ID ou data)
        min_width = 20
        remaining_width = page_width
        
        for i, col in enumerate(df_to_display.columns):
            if 'ID' in col or 'Id' in col or 'id' in col:
                col_widths[i] = min_width
                remaining_width -= min_width
                
        # Cabeçalho
        self.add_table_header(df_to_display.columns, col_widths)
        
        # Adicionar linhas alternando cores
        for i, row in enumerate(df_to_display.itertuples(index=False)):
            self.add_table_row(row, col_widths, fill=(i % 2 == 0))
            
        # Adicionar contador se houver limitação de linhas
        if max_rows is not None and len(df) > max_rows:
            self.ln(5)
            self.add_paragraph(f"Exibindo {max_rows} de {len(df)} registros.")
            
    def add_progress_bar(self, percent, width=180, height=10, label=None):
        # Configurações
        x = self.get_x()
        y = self.get_y()
        
        # Desenhar borda externa
        self.set_draw_color(0, 0, 0)
        self.rect(x, y, width, height)
        
        # Calcular largura preenchida
        filled_width = (percent / 100) * width if percent <= 100 else width
        
        # Desenhar barra de progresso
        if percent < 50:
            self.set_fill_color(231, 76, 60)  # Vermelho para menos de 50%
        elif percent < 75:
            self.set_fill_color(241, 196, 15)  # Amarelo para menos de 75%
        else:
            self.set_fill_color(46, 204, 113)  # Verde para 75% ou mais
            
        if filled_width > 0:
            self.rect(x, y, filled_width, height, 'F')
            
        # Adicionar texto de percentual
        if label:
            self.set_xy(x + width + 5, y)
            self.cell(30, height, label, 0, 1)
        else:
            self.set_xy(x + width + 5, y)
            self.cell(30, height, f"{percent:.1f}%", 0, 1)
            
        # Resetar posição
        self.ln(height + 5)
        
    def add_pie_chart(self, labels, values, title=None, temp_filename='temp_pie.png'):
        """Adiciona um gráfico de pizza ao PDF"""
        if not labels or not values or len(labels) != len(values):
            self.add_paragraph("Dados insuficientes para o gráfico")
            return
            
        # Criar diretório temp se não existir
        temp_dir = os.path.dirname(temp_filename)
        if temp_dir and not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        # Criar o gráfico usando matplotlib
        plt.figure(figsize=(7, 5))
        plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
        plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
        if title:
            plt.title(title)
        plt.tight_layout()
        
        # Salvar o gráfico temporariamente
        plt.savefig(temp_filename)
        plt.close()
        
        # Adicionar o gráfico ao PDF
        if os.path.exists(temp_filename):
            self.image(temp_filename, x=25, w=160)
            os.remove(temp_filename)  # Remover o arquivo temporário
            
    def add_bar_chart(self, labels, values, title=None, temp_filename='temp_bar.png'):
        """Adiciona um gráfico de barras ao PDF"""
        if not labels or not values or len(labels) != len(values):
            self.add_paragraph("Dados insuficientes para o gráfico")
            return
            
        # Criar diretório temp se não existir
        temp_dir = os.path.dirname(temp_filename)
        if temp_dir and not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        # Criar o gráfico usando matplotlib
        plt.figure(figsize=(8, 5))
        plt.bar(labels, values)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        if title:
            plt.title(title)
            
        # Salvar o gráfico temporariamente
        plt.savefig(temp_filename)
        plt.close()
        
        # Adicionar o gráfico ao PDF
        if os.path.exists(temp_filename):
            self.image(temp_filename, x=20, w=170)
            os.remove(temp_filename)  # Remover o arquivo temporário
            
    def add_stacked_bar_chart(self, labels, values_dict, title=None, temp_filename='temp_stacked_bar.png'):
        """Adiciona um gráfico de barras empilhadas ao PDF"""
        if not labels or not values_dict:
            self.add_paragraph("Dados insuficientes para o gráfico")
            return
            
        # Criar diretório temp se não existir
        temp_dir = os.path.dirname(temp_filename)
        if temp_dir and not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        # Criar o gráfico usando matplotlib
        plt.figure(figsize=(8, 5))
        
        bottom = np.zeros(len(labels))
        for label, values in values_dict.items():
            plt.bar(labels, values, bottom=bottom, label=label)
            bottom += np.array(values)
            
        plt.xticks(rotation=45, ha='right')
        plt.legend()
        plt.tight_layout()
        if title:
            plt.title(title)
            
        # Salvar o gráfico temporariamente
        plt.savefig(temp_filename)
        plt.close()
        
        # Adicionar o gráfico ao PDF
        if os.path.exists(temp_filename):
            self.image(temp_filename, x=20, w=170)
            os.remove(temp_filename)  # Remover o arquivo temporário

def generate_commercial_pdf_report(
    output_path,
    mes,
    ano,
    total_meetings,
    new_client_count,
    existing_client_count,
    total_meeting_hours,
    new_client_hours,
    existing_client_hours,
    meta_reunioes,
    percentual_meta,
    meetings_by_user,
    meetings_by_client,
    detailed_meetings
):
    """
    Gera um relatório PDF de indicadores comerciais
    
    Args:
        output_path: Caminho onde o PDF será salvo
        mes, ano: Mês e ano do relatório
        total_meetings, new_client_count, existing_client_count: Contadores de reuniões
        total_meeting_hours, new_client_hours, existing_client_hours: Horas de reuniões
        meta_reunioes, percentual_meta: Meta e percentual de atingimento
        meetings_by_user: DataFrame com reuniões por usuário
        meetings_by_client: DataFrame com reuniões por cliente
        detailed_meetings: DataFrame com detalhamento das reuniões
    
    Returns:
        bool: True se o PDF foi gerado com sucesso, False caso contrário
    """
    try:
        # Criar pasta temporária para os gráficos
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        # Inicializar o PDF
        pdf = CommercialReportPDF(title=f"Relatório de Indicadores Comerciais - {calendar.month_name[mes]} {ano}")
        pdf.set_author("Sistema de Gestão")
        pdf.set_creator("Time Tracker")
        
        # Adicionar a capa
        pdf.add_page()
        
        # Adicionar logo grande na capa
        if os.path.exists('logo.png'):
            pdf.image('logo.png', x=30, y=60, w=150)
            
        # Título principal
        pdf.set_xy(0, 120)
        pdf.set_font('Arial', 'B', 24)
        pdf.cell(0, 20, "Relatório de Indicadores Comerciais", 0, 1, 'C')
        
        # Período do relatório
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, f"{calendar.month_name[mes]} {ano}", 0, 1, 'C')
        
        # Data de geração
        pdf.set_font('Arial', 'I', 12)
        pdf.set_y(-50)
        pdf.cell(0, 10, f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}", 0, 1, 'C')
        
        # Adicionar página de resumo
        pdf.add_page()
        pdf.add_heading("Resumo Executivo")
        
        # Métricas principais em formato de cards
        col_width = 95
        
        # 1ª linha
        pdf.set_font('Arial', 'B', 12)
        
        pdf.set_fill_color(41, 128, 185)  # Azul
        pdf.set_text_color(255, 255, 255)  # Branco
        pdf.rect(pdf.l_margin, pdf.get_y(), col_width, 8, 'F')
        pdf.set_xy(pdf.l_margin, pdf.get_y())
        pdf.cell(col_width, 8, "Total de Reuniões", 0, 0, 'C')
        
        pdf.set_fill_color(46, 204, 113)  # Verde
        pdf.rect(pdf.l_margin + col_width + 5, pdf.get_y(), col_width, 8, 'F')
        pdf.set_xy(pdf.l_margin + col_width + 5, pdf.get_y())
        pdf.cell(col_width, 8, "Total de Horas", 0, 1, 'C')
        
        # Valores da 1ª linha
        pdf.set_text_color(0, 0, 0)  # Preto
        pdf.set_font('Arial', 'B', 20)
        
        pdf.set_xy(pdf.l_margin, pdf.get_y() + 2)
        pdf.cell(col_width, 15, str(total_meetings), 0, 0, 'C')
        
        pdf.set_xy(pdf.l_margin + col_width + 5, pdf.get_y())
        pdf.cell(col_width, 15, format_hours_minutes(total_meeting_hours), 0, 1, 'C')
        
        pdf.ln(5)
        
        # 2ª linha
        pdf.set_font('Arial', 'B', 12)
        
        pdf.set_fill_color(155, 89, 182)  # Roxo
        pdf.set_text_color(255, 255, 255)  # Branco
        pdf.rect(pdf.l_margin, pdf.get_y(), col_width, 8, 'F')
        pdf.set_xy(pdf.l_margin, pdf.get_y())
        pdf.cell(col_width, 8, "Novos Clientes", 0, 0, 'C')
        
        pdf.set_fill_color(231, 76, 60)  # Vermelho
        pdf.rect(pdf.l_margin + col_width + 5, pdf.get_y(), col_width, 8, 'F')
        pdf.set_xy(pdf.l_margin + col_width + 5, pdf.get_y())
        pdf.cell(col_width, 8, "Clientes Existentes", 0, 1, 'C')
        
        # Valores da 2ª linha
        pdf.set_text_color(0, 0, 0)  # Preto
        pdf.set_font('Arial', 'B', 20)
        
        pdf.set_xy(pdf.l_margin, pdf.get_y() + 2)
        pdf.cell(col_width, 15, str(new_client_count), 0, 0, 'C')
        
        pdf.set_xy(pdf.l_margin + col_width + 5, pdf.get_y())
        pdf.cell(col_width, 15, str(existing_client_count), 0, 1, 'C')
        
        pdf.ln(5)
        
        # Meta de reuniões com novos clientes
        pdf.add_subheading("Meta de Reuniões com Novos Clientes")
        
        # Progresso da meta
        pdf.set_font('Arial', '', 10)
        pdf.cell(40, 8, f"Meta: {meta_reunioes} reuniões", 0, 0)
        pdf.cell(40, 8, f"Realizadas: {new_client_count} reuniões", 0, 0)
        pdf.cell(40, 8, f"Progresso: {percentual_meta:.1f}%", 0, 1)
        
        # Barra de progresso
        pdf.add_progress_bar(percentual_meta)
        
        # Preparar dados para gráficos
        # 1. Gráfico de pizza para tipos de reuniões
        if total_meetings > 0:
            pdf.ln(5)
            pdf.add_subheading("Distribuição de Reuniões por Tipo de Cliente")
            
            labels = ["Novos Clientes", "Clientes Existentes"]
            values = [new_client_count, existing_client_count]
            
            pdf.add_pie_chart(
                labels, 
                values, 
                "Distribuição de Reuniões", 
                f"{temp_dir}/pie_meetings.png"
            )
        
        # 2. Gráfico de barras para reuniões por colaborador
        if not meetings_by_user.empty:
            pdf.add_page()
            pdf.add_heading("Análise por Colaborador")
            
            # Ordenar por total de reuniões (decrescente)
            top_users = meetings_by_user.sort_values('total_reunioes', ascending=False).head(10)
            
            # Criar gráfico para Top 10 colaboradores
            labels = top_users['nome_completo'].tolist()
            values = top_users['total_reunioes'].tolist()
            
            pdf.add_bar_chart(
                labels, 
                values, 
                "Top 10 Colaboradores por Número de Reuniões", 
                f"{temp_dir}/bar_users.png"
            )
            
            # Tabela com detalhamento por colaborador
            pdf.ln(5)
            pdf.add_subheading("Detalhamento por Colaborador")
            
            # Formatar a tabela
            table_data = meetings_by_user.copy()
            
            # Renomear colunas para exibição
            table_data = table_data.rename(columns={
                'nome_completo': 'Colaborador',
                'total_reunioes': 'Total Reuniões',
                'total_horas': 'Horas',
                'novos_clientes': 'Novos Clientes'
            })
            
            # Formatar coluna de horas
            table_data['Horas'] = table_data['Horas'].apply(format_hours_minutes)
            
            # Organizar por total de reuniões (decrescente)
            table_data = table_data.sort_values('Total Reuniões', ascending=False)
            
            # Exibir a tabela
            pdf.add_table(table_data)
            
        # 3. Gráfico de reuniões por cliente
        if not meetings_by_client.empty:
            pdf.add_page()
            pdf.add_heading("Análise por Cliente")
            
            # Ordenar por total de reuniões (decrescente)
            top_clients = meetings_by_client.sort_values('total_reunioes', ascending=False).head(10)
            
            # Criar gráfico para Top 10 clientes
            labels = top_clients['client_name'].tolist()
            values = top_clients['total_reunioes'].tolist()
            
            pdf.add_bar_chart(
                labels, 
                values, 
                "Top 10 Clientes por Número de Reuniões", 
                f"{temp_dir}/bar_clients.png"
            )
            
            # Tabela com detalhamento por cliente
            pdf.ln(5)
            pdf.add_subheading("Detalhamento por Cliente")
            
            # Formatar a tabela
            table_data = meetings_by_client.copy()
            
            # Renomear colunas para exibição
            table_data = table_data.rename(columns={
                'client_name': 'Cliente',
                'total_reunioes': 'Total Reuniões',
                'total_horas': 'Horas',
                'tipo_cliente': 'Tipo de Cliente'
            })
            
            # Formatar coluna de horas
            table_data['Horas'] = table_data['Horas'].apply(format_hours_minutes)
            
            # Organizar por total de reuniões (decrescente)
            table_data = table_data.sort_values('Total Reuniões', ascending=False)
            
            # Exibir a tabela
            pdf.add_table(table_data)
        
        # 4. Detalhamento das reuniões
        if not detailed_meetings.empty:
            pdf.add_page()
            pdf.add_heading("Detalhamento das Reuniões")
            
            # Formatar a tabela de detalhamento
            table_data = detailed_meetings.copy()
            
            # Limitar a tabela a 25 registros por página
            max_rows_per_page = 25
            total_rows = len(table_data)
            
            for i in range(0, total_rows, max_rows_per_page):
                if i > 0:
                    pdf.add_page()
                    pdf.add_heading("Detalhamento das Reuniões (continuação)")
                
                # Obter a parte da tabela para esta página
                part = table_data.iloc[i:min(i + max_rows_per_page, total_rows)]
                
                # Exibir a parte da tabela
                pdf.add_table(part)
                
                # Adicionar contador
                pdf.set_font('Arial', 'I', 8)
                pdf.cell(0, 5, f"Exibindo registros {i+1} a {min(i + max_rows_per_page, total_rows)} de {total_rows}", 0, 1, 'R')
            
        # Salvar o PDF
        pdf.output(output_path)
        
        # Remover diretório temporário se estiver vazio
        if os.path.exists(temp_dir) and not os.listdir(temp_dir):
            os.rmdir(temp_dir)
            
        return True
        
    except Exception as e:
        st.error(f"Erro ao gerar PDF: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return False 
    
    #Formatar a tabela de detalhamento
        table_data = detailed_meetings.copy()
            
        # Aplicar as formatações necessárias
        # Formatar data
        if 'Data' in table_data.columns:
            table_data['Data'] = pd.to_datetime(table_data['Data']).dt.strftime('%d/%m/%Y')
        
        # Limitar a tabela a 25 registros por página
        max_rows_per_page = 25
        total_rows = len(table_data)
        
        for i in range(0, total_rows, max_rows_per_page):
            if i > 0:
                pdf.add_page()
                pdf.add_heading("Detalhamento das Reuniões (continuação)")
            
            # Obter a parte da tabela para esta página
            part = table_data.iloc[i:min(i + max_rows_per_page, total_rows)]
            
            # Exibir a parte da tabela
            pdf.add_table(part)
            
            # Adicionar contador
            pdf.set_font('Arial', 'I', 8)
            pdf.cell(0, 5, f"Exibindo registros {i+1} a {min(i + max_rows_per_page, total_rows)} de {total_rows}", 0, 1, 'R')
        
        # Salvar o PDF
        pdf.output(output_path)
        
        # Remover diretório temporário se estiver vazio
        if os.path.exists(temp_dir) and not os.listdir(temp_dir):
            os.rmdir(temp_dir)
            
        return True
    
    except Exception as e:
        st.error(f"Erro ao gerar PDF: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return False

def commercial_indicators_email():
    """Relatório por Email de Indicadores Comerciais"""
    st.title("📧 Email de Indicadores Comerciais")
    
    # Verificar o papel do usuário
    is_admin = st.session_state.user_info['role'].lower() == 'admin'
    if not is_admin:
        st.warning("Esta funcionalidade é exclusiva para administradores.")
        return
    
    # Inicializar o DatabaseManager
    db_manager = DatabaseManager()
    
    try:
        # Carregar tabelas necessárias
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        groups_df = db_manager.query_to_df("SELECT * FROM groups")
        clients_df = db_manager.query_to_df("SELECT * FROM clients")
        
        # Tentar carregar categorias e atividades, com tratamento de erro
        try:
            categories_df = db_manager.query_to_df("SELECT * FROM task_categories")
        except:
            categories_df = pd.DataFrame(columns=['task_category_id', 'task_category'])
            st.warning("Tabela de categorias não encontrada.")
            
        try:
            activities_df = db_manager.query_to_df("SELECT * FROM activities")
        except:
            activities_df = pd.DataFrame(columns=['activity_id', 'activity_name'])
            st.warning("Tabela de atividades não encontrada.")
        
        # Carregar dados de timesheet
        timesheet_df = db_manager.query_to_df("SELECT * FROM timesheet")
        
        # Converter datas para datetime
        if not timesheet_df.empty:
            timesheet_df['start_date'] = pd.to_datetime(timesheet_df['start_date'], format='mixed', errors='coerce')
        
        # Verificar se existe a coluna new_client na tabela timesheet
        if 'new_client' not in timesheet_df.columns:
            timesheet_df['new_client'] = 0
            st.warning("Coluna 'new_client' não encontrada na tabela timesheet. Usando valor padrão 0.")
        
        # Interface para configuração do email
        st.subheader("Configuração do Email")
        
        with st.form("email_config_form"):
            # Seleção de período
            col1, col2 = st.columns(2)
            
            with col1:
                current_month = datetime.now().month
                mes = st.selectbox(
                    "Mês",
                    options=list(range(1, 13)),
                    format_func=lambda x: calendar.month_name[x],
                    index=current_month - 1
                )
            
            with col2:
                current_year = datetime.now().year
                ano = st.selectbox(
                    "Ano",
                    options=list(range(current_year - 3, current_year + 1)),
                    index=3
                )
            
            # Formato do relatório
            report_format = st.radio(
                "Formato do Relatório",
                ["PDF", "Excel", "PDF e Excel"],
                index=2
            )
            
            # Destinatários do email
            recipients = st.text_input(
                "Destinatários (separados por vírgula)",
                help="Email dos destinatários separados por vírgula"
            )
            
            # Assunto do email
            subject = st.text_input(
                "Assunto do Email", 
                f"Relatório de Indicadores Comerciais - {calendar.month_name[mes]} {ano}"
            )
            
            # Conteúdo do email
            email_message = st.text_area(
                "Mensagem do Email",
                f"""Prezados,

Em anexo, relatório de indicadores comerciais referente ao mês de {calendar.month_name[mes]} de {ano}.

Atenciosamente,
Equipe de Gestão"""
            )
            
            # SMTP settings
            with st.expander("Configurações de SMTP"):
                smtp_server = st.text_input("Servidor SMTP", "smtp.office365.com")
                smtp_port = st.number_input("Porta SMTP", value=587, step=1)
                smtp_user = st.text_input("Usuário SMTP", "notifications@grupoerre.pt")
                smtp_password = st.text_input("Senha SMTP", type="password", value="9FWkMpK8tif2lY4")
                use_tls = st.checkbox("Usar TLS", value=True)
            
            # Botão para gerar e enviar o email
            submit_button = st.form_submit_button("Gerar e Enviar Relatório")
        
        if submit_button:
            # Validar entradas
            if not recipients and not st.checkbox("Apenas gerar relatório sem enviar", value=False):
                st.error("Por favor, informe pelo menos um destinatário ou marque a opção para apenas gerar o relatório.")
                return
            
            with st.spinner("Gerando relatório..."):
                # Definir início e fim do mês selecionado
                inicio_mes = datetime(ano, mes, 1)
                ultimo_dia = calendar.monthrange(ano, mes)[1]
                fim_mes = datetime(ano, mes, ultimo_dia, 23, 59, 59)
                
                # Filtrar dados da equipe comercial (group_id 5)
                # Primeiro identificar os usuários da equipe comercial
                commercial_team_users = []
                for _, user in users_df.iterrows():
                    try:
                        if isinstance(user['groups'], str):
                            user_groups = eval(user['groups'])
                            if isinstance(user_groups, dict):
                                user_groups = list(user_groups.values())
                            
                            # Verificar se o usuário pertence ao grupo comercial
                            commercial_group = groups_df[groups_df['id'] == 5]['group_name'].iloc[0] if not groups_df[groups_df['id'] == 5].empty else None
                            
                            if commercial_group and commercial_group in user_groups:
                                commercial_team_users.append(user['user_id'])
                    except Exception as e:
                        continue
                
                # Filtrar registros de timesheet da equipe comercial
                if commercial_team_users:
                    user_ids_str = ','.join(str(uid) for uid in commercial_team_users)
                    commercial_data = timesheet_df[
                        (timesheet_df['user_id'].astype(str).isin([str(uid) for uid in commercial_team_users])) &
                        (timesheet_df['start_date'] >= inicio_mes) &
                        (timesheet_df['start_date'] <= fim_mes)
                    ]
                else:
                    st.warning("Não foram encontrados usuários na equipe comercial (group_id 5).")
                    commercial_data = pd.DataFrame()
                
                # Filtrar apenas reuniões (task_category ou activity relacionada a reuniões)
                meeting_data = pd.DataFrame()
                
                if not commercial_data.empty:
                    # Buscar categorias e atividades relacionadas a reuniões
                    # Verificar se existem colunas de categoria e atividade
                    has_category = 'category_id' in commercial_data.columns
                    has_activity = 'task_id' in commercial_data.columns
                    
                    # Filtragem baseada em categorias
                    if has_category:
                        # Juntar com informações de categoria
                        meeting_data_by_category = commercial_data.merge(
                            categories_df[['task_category_id', 'task_category']],
                            left_on='category_id',
                            right_on='task_category_id',
                            how='left'
                        )
                        
                        # Filtrar apenas reuniões
                        meetings_category = meeting_data_by_category[
                            meeting_data_by_category['task_category'].str.lower().str.contains('reuni', na=False) |
                            meeting_data_by_category['task_category'].str.lower().str.contains('meet', na=False)
                        ] if not meeting_data_by_category.empty else pd.DataFrame()
                        
                        # Adicionar aos dados de reunião
                        meeting_data = pd.concat([meeting_data, meetings_category])
                    
                    # Filtragem baseada em atividades
                    if has_activity:
                        # Juntar com informações de atividade
                        meeting_data_by_activity = commercial_data.merge(
                            activities_df[['activity_id', 'activity_name']],
                            left_on='task_id',
                            right_on='activity_id',
                            how='left'
                        )
                        
                        # Filtrar apenas reuniões
                        meetings_activity = meeting_data_by_activity[
                            meeting_data_by_activity['activity_name'].str.lower().str.contains('reuni', na=False) |
                            meeting_data_by_activity['activity_name'].str.lower().str.contains('meet', na=False)
                        ] if not meeting_data_by_activity.empty else pd.DataFrame()
                        
                        # Adicionar aos dados de reunião (evitando duplicatas)
                        if not meeting_data.empty:
                            # Obter IDs das reuniões já incluídas
                            existing_ids = meeting_data['id'].astype(str).tolist() if 'id' in meeting_data.columns else []
                            
                            # Adicionar apenas reuniões não incluídas anteriormente
                            new_meetings = meetings_activity[~meetings_activity['id'].astype(str).isin(existing_ids)] if not meetings_activity.empty else pd.DataFrame()
                            meeting_data = pd.concat([meeting_data, new_meetings])
                        else:
                            meeting_data = meetings_activity
                    
                    # Se não houver categorias ou atividades específicas, considerar todos os registros
                    if (not has_category and not has_activity) or meeting_data.empty:
                        # Se não encontramos nada pelos métodos acima, vamos buscar pela descrição
                        meeting_data = commercial_data[
                            commercial_data['description'].str.lower().str.contains('reuni', na=False) |
                            commercial_data['description'].str.lower().str.contains('meet', na=False)
                        ] if not commercial_data.empty else pd.DataFrame()
                
                # Analisar reuniões com novos clientes
                if not meeting_data.empty:
                    # Separar reuniões por tipo de cliente (novo ou existente)
                    new_client_meetings = meeting_data[meeting_data['new_client'] == 1]
                    existing_client_meetings = meeting_data[meeting_data['new_client'] == 0]
                    
                    # Juntar com informações de usuários e clientes
                    enriched_meetings = meeting_data.merge(
                        users_df[['user_id', 'First_Name', 'Last_Name']],
                        on='user_id',
                        how='left'
                    )
                    
                    # Juntar com clientes se client_id existir
                    if 'client_id' in enriched_meetings.columns:
                        enriched_meetings = enriched_meetings.merge(
                            clients_df[['client_id', 'name']].rename(columns={'name': 'client_name'}),
                            on='client_id',
                            how='left'
                        )
                    else:
                        enriched_meetings['client_name'] = "Cliente não especificado"
                    
                    # Adicionar coluna de nome completo
                    enriched_meetings['nome_completo'] = enriched_meetings['First_Name'] + ' ' + enriched_meetings['Last_Name']
                    
                    # Juntar com categorias e atividades se existirem
                    if 'category_id' in enriched_meetings.columns:
                        enriched_meetings = enriched_meetings.merge(
                            categories_df[['task_category_id', 'task_category']],
                            left_on='category_id',
                            right_on='task_category_id',
                            how='left'
                        )
                    
                    if 'task_id' in enriched_meetings.columns:
                        enriched_meetings = enriched_meetings.merge(
                            activities_df[['activity_id', 'activity_name']],
                            left_on='task_id',
                            right_on='activity_id',
                            how='left'
                        )
                    
                    # Calcular métricas
                    total_meetings = len(meeting_data)
                    new_client_count = len(new_client_meetings)
                    existing_client_count = len(existing_client_meetings)
                    
                    # Calcular total de horas em reuniões
                    total_meeting_hours = meeting_data['hours'].sum()
                    new_client_hours = new_client_meetings['hours'].sum() if not new_client_meetings.empty else 0
                    existing_client_hours = existing_client_meetings['hours'].sum() if not existing_client_meetings.empty else 0
                    
                    # Calcular percentual de atingimento da meta (15 reuniões com novos clientes)
                    meta_reunioes = 20
                    percentual_meta = (new_client_count / meta_reunioes * 100) if meta_reunioes > 0 else 0
                    
                    # Reuniões por colaborador
                    meetings_by_user = enriched_meetings.groupby('nome_completo').agg({
                        'id': 'count',
                        'hours': 'sum',
                        'new_client': lambda x: sum(x == 1)  # Contar apenas reuniões com novos clientes
                    }).reset_index()
                    
                    meetings_by_user.rename(columns={
                        'id': 'total_reunioes',
                        'hours': 'total_horas',
                        'new_client': 'novos_clientes'
                    }, inplace=True)
                    
                    # Reuniões por cliente
                    meetings_by_client = enriched_meetings.groupby('client_name').agg({
                        'id': 'count',
                        'hours': 'sum',
                        'new_client': 'first'  # Pegar o primeiro valor (todos serão iguais para o mesmo cliente)
                    }).reset_index()
                    
                    meetings_by_client.rename(columns={
                        'id': 'total_reunioes',
                        'hours': 'total_horas',
                        'new_client': 'cliente_novo'
                    }, inplace=True)
                    
                    # Converter o indicador para texto
                    meetings_by_client['tipo_cliente'] = meetings_by_client['cliente_novo'].apply(
                        lambda x: "Novo Cliente" if x == 1 else "Cliente Existente"
                    )
                    
                    # Preparar dados para o detalhamento
                    detailed_meetings = enriched_meetings.copy()
                    
                    # Selecionar colunas relevantes
                    detail_columns = [
                        'nome_completo', 'client_name', 'start_date', 'hours', 
                        'description', 'new_client'
                    ]
                    
                    # Adicionar colunas de categoria e atividade se existirem
                    if 'task_category' in detailed_meetings.columns:
                        detail_columns.append('task_category')
                    if 'activity_name' in detailed_meetings.columns:
                        detail_columns.append('activity_name')
                    
                    # Filtrar apenas colunas existentes
                    available_columns = [col for col in detail_columns if col in detailed_meetings.columns]
                    
                    # Preparar dados para o detalhamento
                    detail_data = detailed_meetings[available_columns].copy()
                    
                    # Formatar colunas
                    if 'start_date' in detail_data.columns:
                        detail_data['start_date'] = detail_data['start_date'].dt.strftime('%d/%m/%Y %H:%M')
                    if 'hours' in detail_data.columns:
                        detail_data['hours_formatted'] = detail_data['hours'].apply(format_hours_minutes)
                    if 'new_client' in detail_data.columns:
                        detail_data['new_client_text'] = detail_data['new_client'].apply(
                            lambda x: "Sim" if x == 1 else "Não"
                        )
                    
                    # Renomear colunas para o relatório
                    column_rename = {
                        'nome_completo': 'Colaborador',
                        'client_name': 'Cliente',
                        'start_date': 'Data',
                        'hours': 'Horas',
                        'hours_formatted': 'Horas (HH:MM)',
                        'description': 'Descrição',
                        'new_client': 'Novo Cliente',
                        'new_client_text': 'Novo Cliente',
                        'task_category': 'Categoria',
                        'activity_name': 'Atividade'
                    }
                    
                    # Filtrar apenas colunas existentes
                    rename_map = {k: v for k, v in column_rename.items() if k in detail_data.columns}
                    
                    # Aplicar renomeação
                    detail_data = detail_data.rename(columns=rename_map)
                    
                    # Criar diretório temporário para os arquivos
                    temp_dir = tempfile.mkdtemp()
                    
                    # Definir caminhos para os arquivos
                    excel_path = None
                    pdf_path = None
                    
                    # Gerar relatório Excel se solicitado
                    if "Excel" in report_format:
                        excel_path = os.path.join(temp_dir, f"indicadores_comerciais_{mes:02d}_{ano}.xlsx")
                        
                        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                            # Resumo
                            resumo_df = pd.DataFrame([{
                                'Mês/Ano': f"{calendar.month_name[mes]} {ano}",
                                'Total de Reuniões': total_meetings,
                                'Reuniões com Novos Clientes': new_client_count,
                                'Reuniões com Clientes Existentes': existing_client_count,
                                'Total de Horas em Reuniões': total_meeting_hours,
                                'Horas com Novos Clientes': new_client_hours,
                                'Horas com Clientes Existentes': existing_client_hours,
                                'Meta de Reuniões com Novos Clientes': meta_reunioes,
                                'Percentual da Meta Atingido': percentual_meta
                            }])
                            
                            resumo_df.to_excel(writer, sheet_name='Resumo', index=False)
                            
                            # Reuniões por colaborador
                            if not meetings_by_user.empty:
                                meetings_by_user.to_excel(writer, sheet_name='Por Colaborador', index=False)
                            
                            # Reuniões por cliente
                            if not meetings_by_client.empty:
                                meetings_by_client.to_excel(writer, sheet_name='Por Cliente', index=False)
                            
                            # Detalhamento completo
                            if not detail_data.empty:
                                detail_data.to_excel(writer, sheet_name='Detalhamento', index=False)
                                
                        st.success(f"Relatório Excel gerado com sucesso: {excel_path}")
                    
                    # Gerar relatório PDF se solicitado
                    if "PDF" in report_format:
                        pdf_path = os.path.join(temp_dir, f"indicadores_comerciais_{mes:02d}_{ano}.pdf")
                        
                        # Gerar PDF usando a função de relatório
                        pdf_success = generate_commercial_pdf_report(
                            pdf_path,
                            mes,
                            ano,
                            total_meetings,
                            new_client_count,
                            existing_client_count,
                            total_meeting_hours,
                            new_client_hours,
                            existing_client_hours,
                            meta_reunioes,
                            percentual_meta,
                            meetings_by_user,
                            meetings_by_client,
                            detail_data
                        )
                        
                        if pdf_success:
                            st.success(f"Relatório PDF gerado com sucesso: {pdf_path}")
                        else:
                            st.error("Falha ao gerar o relatório PDF.")
                    
                    # Enviar email se houver destinatários
                    if recipients:
                        try:
                            # Criar email
                            msg = MIMEMultipart()
                            msg['From'] = smtp_user
                            msg['To'] = recipients
                            msg['Subject'] = subject
                            
                            # Adicionar corpo do email
                            msg.attach(MIMEText(email_message, 'plain'))
                            
                            # Adicionar anexos
                            if excel_path and os.path.exists(excel_path):
                                with open(excel_path, 'rb') as f:
                                    attachment = MIMEApplication(f.read())
                                    attachment.add_header(
                                        'Content-Disposition', 
                                        'attachment', 
                                        filename=os.path.basename(excel_path)
                                    )
                                    msg.attach(attachment)
                            
                            if pdf_path and os.path.exists(pdf_path):
                                with open(pdf_path, 'rb') as f:
                                    attachment = MIMEApplication(f.read())
                                    attachment.add_header(
                                        'Content-Disposition', 
                                        'attachment', 
                                        filename=os.path.basename(pdf_path)
                                    )
                                    msg.attach(attachment)
                            
                            # Configurar e enviar email
                            server = smtplib.SMTP(smtp_server, smtp_port)
                            
                            if use_tls:
                                server.starttls()
                            
                            if smtp_user and smtp_password:
                                server.login(smtp_user, smtp_password)
                            
                            server.send_message(msg)
                            server.quit()
                            
                            st.success(f"Email enviado com sucesso para {recipients}!")
                            
                        except Exception as e:
                            st.error(f"Erro ao enviar email: {str(e)}")
                    
                    # Oferecer download dos relatórios
                    col1, col2 = st.columns(2)
                    
                    if excel_path and os.path.exists(excel_path):
                        with open(excel_path, 'rb') as f:
                            col1.download_button(
                                label="📥 Baixar Relatório Excel",
                                data=f,
                                file_name=os.path.basename(excel_path),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    
                    if pdf_path and os.path.exists(pdf_path):
                        with open(pdf_path, 'rb') as f:
                            col2.download_button(
                                label="📥 Baixar Relatório PDF",
                                data=f,
                                file_name=os.path.basename(pdf_path),
                                mime="application/pdf"
                            )
                    
                else:
                    st.warning("Não foram encontradas reuniões para a equipe comercial no período selecionado.")
                    
                    # Criar relatórios vazios
                    temp_dir = tempfile.mkdtemp()
                    
                    # Definir variáveis padrão
                    total_meetings = 0
                    new_client_count = 0
                    existing_client_count = 0
                    total_meeting_hours = 0
                    new_client_hours = 0
                    existing_client_hours = 0
                    meta_reunioes = 15
                    percentual_meta = 0
                    meetings_by_user = pd.DataFrame(columns=['nome_completo', 'total_reunioes', 'total_horas', 'novos_clientes'])
                    meetings_by_client = pd.DataFrame(columns=['client_name', 'total_reunioes', 'total_horas', 'tipo_cliente'])
                    detail_data = pd.DataFrame(columns=['Colaborador', 'Cliente', 'Data', 'Horas', 'Descrição', 'Novo Cliente'])
                    
                    # Criar Excel vazio
                    if "Excel" in report_format:
                        excel_path = os.path.join(temp_dir, f"indicadores_comerciais_vazio_{mes:02d}_{ano}.xlsx")
                        
                        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                            empty_df = pd.DataFrame([{
                                'Mês/Ano': f"{calendar.month_name[mes]} {ano}",
                                'Situação': "Não foram encontradas reuniões comerciais no período",
                                'Meta de Reuniões com Novos Clientes': 15,
                                'Reuniões Realizadas': 0,
                                'Percentual da Meta': 0,
                            }])
                            
                            empty_df.to_excel(writer, sheet_name='Resumo', index=False)
                        
                        st.download_button(
                            label="📥 Baixar Relatório Excel Vazio",
                            data=open(excel_path, 'rb'),
                            file_name=os.path.basename(excel_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    # Criar PDF vazio
                    if "PDF" in report_format:
                        pdf_path = os.path.join(temp_dir, f"indicadores_comerciais_vazio_{mes:02d}_{ano}.pdf")
                        
                        # Gerar PDF usando a função de relatório
                        pdf_success = generate_commercial_pdf_report(
                            pdf_path,
                            mes,
                            ano,
                            total_meetings,
                            new_client_count,
                            existing_client_count,
                            total_meeting_hours,
                            new_client_hours,
                            existing_client_hours,
                            meta_reunioes,
                            percentual_meta,
                            meetings_by_user,
                            meetings_by_client,
                            detail_data
                        )
                        
                        if pdf_success:
                            st.download_button(
                                label="📥 Baixar Relatório PDF Vazio",
                                data=open(pdf_path, 'rb'),
                                file_name=os.path.basename(pdf_path),
                                mime="application/pdf"
                            )
                    
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {str(e)}")
        import traceback
        st.error(traceback.format_exc())

if __name__ == "__main__":
    commercial_indicators_email()