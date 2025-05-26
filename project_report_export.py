import os
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from datetime import datetime
import calendar
from fpdf import FPDF
import io
import base64
import streamlit as st
from report_utils import calcular_dias_uteis_projeto

# Definir a função format_hours_minutes aqui em vez de importá-la
def format_hours_minutes(hours):
    """Converte horas decimais para formato HH:mm"""
    if pd.isna(hours) or not isinstance(hours, (int, float)):
        return "00:00"
    total_minutes = int(hours * 60)
    hours_part = total_minutes // 60
    minutes_part = total_minutes % 60
    return f"{hours_part:02d}:{minutes_part:02d}"

def generate_single_project_pdf(
    projeto_info, 
    metricas, 
    entries, 
    recursos_utilizacao, 
    horas_por_categoria=None, 
    horas_por_atividade=None, 
    consumo_mensal=None, 
    phases_info=None, 
    users_df=None,
    rates_df=None,
    client_name=None
):
    """
    Gera um relatório PDF para um único projeto
    
    Args:
        projeto_info: Informações básicas do projeto
        metricas: Métricas calculadas do projeto
        entries: DataFrame com registros de horas
        recursos_utilizacao: DataFrame com dados de utilização por recurso
        horas_por_categoria: DataFrame com horas por categoria (opcional)
        horas_por_atividade: DataFrame com horas por atividade (opcional)
        consumo_mensal: DataFrame com consumo mensal (opcional)
        phases_info: Informações das fases do projeto (opcional)
        users_df: DataFrame com informações dos usuários
        rates_df: DataFrame com informações de rates
        client_name: Nome do cliente
        
    Returns:
        bytes: PDF em formato de bytes para download
    """
    try:
        # Criar diretório temporário para gráficos
        temp_dir = tempfile.mkdtemp()
        
        # Configurar classe PDF customizada
        class PDF(FPDF):
            def header(self):
                # Logo apenas a partir da segunda página
                if self.page_no() != 1:  # Não mostrar o logo na primeira página (capa)
                    if os.path.exists('logo.png'):
                        self.image('logo.png', 10, 8, 33)
                    
                    # Título
                    self.set_font('Arial', 'B', 15)
                    self.cell(80)
                    self.cell(30, 25, 'Relatório de Projeto', 0, 0, 'C')
                    self.ln(20)
            
            def footer(self):
                # Posicionar a 1.5 cm do final
                self.set_y(-15)
                # Fonte Arial itálico 8
                self.set_font('Arial', 'I', 8)
                # Número da página
                self.cell(0, 10, f'Página {self.page_no()}/{{nb}}', 0, 0, 'C')
        
        pdf = PDF()
        pdf.alias_nb_pages()
        
        # Capa do relatório
        pdf.add_page()

        # Adicionar o logo apenas uma vez na capa
        if os.path.exists('logo.png'):
            pdf.image('logo.png', x=10, y=20, w=60)

        pdf.set_font('Arial', 'B', 24)
        pdf.set_xy(0, 80)
        pdf.cell(210, 20, 'Relatório do Projeto', 0, 1, 'C')
        
        # Nome do projeto em tamanho maior
        pdf.set_font('Arial', 'B', 14)
        pdf.set_xy(0, 100)
        pdf.cell(210, 20, projeto_info['project_name'], 0, 1, 'C')
        
        # Cliente
        pdf.set_font('Arial', 'B', 16)
        pdf.set_xy(0, 125)
        pdf.cell(210, 10, client_name, 0, 1, 'C')
        
        # Tipo de projeto
        pdf.set_font('Arial', '', 14)
        pdf.set_xy(0, 140)
        pdf.cell(210, 10, f"Tipo: {projeto_info['project_type']}", 0, 1, 'C')
        
        # Período do projeto
        start_date = pd.to_datetime(projeto_info['start_date']).strftime('%d/%m/%Y')
        end_date = pd.to_datetime(projeto_info['end_date']).strftime('%d/%m/%Y')
        
        pdf.set_font('Arial', '', 14)
        pdf.set_xy(0, 155)
        pdf.cell(210, 10, f"Período: {start_date} a {end_date}", 0, 1, 'C')
        
        # Data de geração
        pdf.set_font('Arial', 'I', 10)
        pdf.set_y(-30)
        pdf.cell(0, 10, f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}", 0, 1, 'C')
        
        # 1. Informações Gerais do Projeto
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, '1. Informações Gerais', 0, 1)
        
        pdf.ln(5)
        pdf.set_font('Arial', '', 11)
        
        # Informações básicas em formato de tabela
        info_width = 60
        value_width = 130
        line_height = 8
        
        # Cliente
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(info_width, line_height, 'Cliente:', 0, 0)
        pdf.set_font('Arial', '', 11)
        pdf.cell(value_width, line_height, client_name, 0, 1)
        
        # Projeto
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(info_width, line_height, 'Projeto:', 0, 0)
        pdf.set_font('Arial', '', 11)
        pdf.cell(value_width, line_height, projeto_info['project_name'], 0, 1)
        
        # Tipo
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(info_width, line_height, 'Tipo:', 0, 0)
        pdf.set_font('Arial', '', 11)
        pdf.cell(value_width, line_height, projeto_info['project_type'], 0, 1)
        
        # Datas
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(info_width, line_height, 'Data de Início:', 0, 0)
        pdf.set_font('Arial', '', 11)
        pdf.cell(value_width, line_height, start_date, 0, 1)
        
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(info_width, line_height, 'Data de Término:', 0, 0)
        pdf.set_font('Arial', '', 11)
        pdf.cell(value_width, line_height, end_date, 0, 1)
        
        # Status
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(info_width, line_height, 'Status:', 0, 0)
        pdf.set_font('Arial', '', 11)
        pdf.cell(value_width, line_height, projeto_info['status'].title(), 0, 1)
        
        pdf.ln(5)
        
        # Seção de métricas principais
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Métricas Principais:', 0, 1)
        
        # Quadro de métricas
        metrics_line_height = 10
        
        # Tabela com 2 colunas de métricas
        col_width = 95
        
        # Linha 1
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(col_width, metrics_line_height, 'Custo Total Planejado:', 0, 0)
        pdf.cell(col_width, metrics_line_height, 'Custo Realizado:', 0, 1)
        
        pdf.set_font('Arial', '', 11)
        pdf.cell(col_width, metrics_line_height, f"EUR {projeto_info['total_cost']:,.2f}", 0, 0)
        pdf.cell(col_width, metrics_line_height, f"EUR {metricas['custo_realizado']:,.2f}", 0, 1)
        
        # Linha 2
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(col_width, metrics_line_height, 'Horas Planejadas:', 0, 0)
        pdf.cell(col_width, metrics_line_height, 'Horas Realizadas:', 0, 1)
        
        pdf.set_font('Arial', '', 11)
        horas_planejadas_fmt = format_hours_minutes(projeto_info['total_hours'])
        horas_realizadas_fmt = format_hours_minutes(metricas['horas_realizadas'])
        
        pdf.cell(col_width, metrics_line_height, horas_planejadas_fmt, 0, 0)
        pdf.cell(col_width, metrics_line_height, horas_realizadas_fmt, 0, 1)
        
        # Linha 3
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(col_width, metrics_line_height, 'Horas Disponíveis:', 0, 0)
        pdf.cell(col_width, metrics_line_height, 'Projeção Final (EAC):', 0, 1)
        
        pdf.set_font('Arial', '', 11)
        horas_disponiveis = max(0, projeto_info['total_hours'] - metricas['horas_realizadas'])
        horas_disponiveis_fmt = format_hours_minutes(horas_disponiveis)
        
        pdf.cell(col_width, metrics_line_height, horas_disponiveis_fmt, 0, 0)
        pdf.cell(col_width, metrics_line_height, f"EUR {metricas['eac']:,.2f}", 0, 1)
        
        # Linha 4
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(col_width, metrics_line_height, 'CPI:', 0, 0)
        pdf.cell(col_width, metrics_line_height, '% Horas Consumidas:', 0, 1)
        
        pdf.set_font('Arial', '', 11)
        # Definir cor do CPI com base no valor
        if metricas['cpi'] >= 1:
            pdf.set_text_color(0, 128, 0)  # Verde
        else:
            pdf.set_text_color(255, 0, 0)  # Vermelho
            
        pdf.cell(col_width, metrics_line_height, f"{metricas['cpi']:.2f}", 0, 0)
        
        # Resetar cor para preto
        pdf.set_text_color(0, 0, 0)
        
        perc_horas = (metricas['horas_realizadas'] / projeto_info['total_hours'] * 100) if projeto_info['total_hours'] > 0 else 0
        pdf.cell(col_width, metrics_line_height, f"{perc_horas:.1f}%", 0, 1)
        
        pdf.ln(5)
        
        # Legenda do CPI
        pdf.set_font('Arial', 'I', 10)
        pdf.multi_cell(0, 5, "CPI (Cost Performance Index): Razão entre o valor planejado e o custo real. Um CPI maior que 1 indica que o projeto está gastando menos do que o planejado (positivo), enquanto um CPI menor que 1 indica que está gastando mais do que o planejado (negativo).")
        
        # 2. Análise de Recursos
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, '2. Utilização de Recursos', 0, 1)
        
        # Verificar se temos dados de recursos
        if recursos_utilizacao is not None and not recursos_utilizacao.empty:
            # Calcular totais
            total_horas = recursos_utilizacao['hours'].sum()
            total_custo = recursos_utilizacao['custo'].sum() if 'custo' in recursos_utilizacao.columns else 0
            
            # Adicionar aviso sobre horas extras
            if metricas['horas_extras_originais'] > 0:
                pdf.set_font('Arial', 'B', 10)
                pdf.cell(0, 10, f"Atenção: Este projeto possui {format_hours_minutes(metricas['horas_extras_originais'])} de horas extras, que são contabilizadas em dobro.", 0, 1)
            
            # Criar tabela de utilização por recurso
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Detalhamento por Recurso:', 0, 1)
            
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 10)
            
            # Definir larguras das colunas
            col_widths = [60, 30, 40, 30, 30]
            
            pdf.cell(col_widths[0], 8, 'Colaborador', 1, 0, 'C', True)
            pdf.cell(col_widths[1], 8, 'Horas', 1, 0, 'C', True)
            pdf.cell(col_widths[2], 8, 'Custo', 1, 0, 'C', True)
            pdf.cell(col_widths[3], 8, '% Horas', 1, 0, 'C', True)
            pdf.cell(col_widths[4], 8, '% Custo', 1, 1, 'C', True)
            
            # Conteúdo da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 9)
            
            # Ordenar por horas (decrescente)
            recursos_ordenados = recursos_utilizacao.sort_values('hours', ascending=False)
            
            # Adicionar linhas da tabela
            row_count = 0
            for _, row in recursos_ordenados.iterrows():
                # Alternar cores das linhas
                if row_count % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)  # Branco
                    fill = False
                else:
                    pdf.set_fill_color(240, 240, 240)  # Cinza claro
                    fill = True
                
                # Nome do colaborador
                pdf.cell(col_widths[0], 7, str(row['nome_completo']), 1, 0, 'L', fill)
                
                # Horas em formato HH:MM
                horas_fmt = format_hours_minutes(row['hours'])
                pdf.cell(col_widths[1], 7, horas_fmt, 1, 0, 'C', fill)
                
                # Custo
                custo = row['custo'] if 'custo' in row else 0
                pdf.cell(col_widths[2], 7, f"EUR {custo:.2f}", 1, 0, 'R', fill)
                
                # Percentual de horas
                perc_horas = row['perc_horas'] if 'perc_horas' in row else (row['hours'] / total_horas * 100 if total_horas > 0 else 0)
                pdf.cell(col_widths[3], 7, f"{perc_horas:.1f}%", 1, 0, 'C', fill)
                
                # Percentual de custo
                perc_custo = row['perc_custo'] if 'perc_custo' in row else (custo / total_custo * 100 if total_custo > 0 else 0)
                pdf.cell(col_widths[4], 7, f"{perc_custo:.1f}%", 1, 1, 'C', fill)
                
                row_count += 1
            
            # Adicionar linha de totais
            pdf.set_font('Arial', 'B', 10)
            pdf.set_fill_color(220, 220, 220)  # Cinza mais escuro para totais
            pdf.cell(col_widths[0], 8, 'TOTAL', 1, 0, 'L', True)
            pdf.cell(col_widths[1], 8, format_hours_minutes(total_horas), 1, 0, 'C', True)
            pdf.cell(col_widths[2], 8, f"EUR {total_custo:.2f}", 1, 0, 'R', True)
            pdf.cell(col_widths[3], 8, '100.0%', 1, 0, 'C', True)
            pdf.cell(col_widths[4], 8, '100.0%', 1, 1, 'C', True)
            
            # Criar gráficos de recursos
            # Gráfico de distribuição de horas por recurso
            try:
                plt.figure(figsize=(10, 6))
                plt.bar(
                    recursos_ordenados['nome_completo'], 
                    recursos_ordenados['hours'],
                    color='#1E88E5'
                )
                plt.title('Distribuição de Horas por Recurso')
                plt.xlabel('Colaborador')
                plt.ylabel('Horas Trabalhadas')
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                
                # Salvar o gráfico em um arquivo temporário
                chart_path = os.path.join(temp_dir, 'recursos_horas.png')
                plt.savefig(chart_path)
                plt.close()
                
                # Adicionar o gráfico ao PDF
                pdf.ln(10)
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(0, 10, 'Gráfico de Distribuição de Horas:', 0, 1)
                pdf.image(chart_path, x=10, y=pdf.get_y(), w=190)
                
                # Ajustar posição Y após o gráfico
                pdf.set_y(pdf.get_y() + 130)  # Ajuste aproximado para altura do gráfico
            except Exception as e:
                pdf.ln(3)
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 10, f"Não foi possível gerar o gráfico de horas: {str(e)}", 0, 1)
        else:
            pdf.set_font('Arial', '', 11)
            pdf.cell(0, 10, "Não há dados de utilização de recursos disponíveis.", 0, 1)
        
        # 3. Consumo Mensal
        if consumo_mensal is not None and not consumo_mensal.empty:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, '3. Consumo de Horas Mensais', 0, 1)
            
            # Resumo do consumo
            pdf.ln(3)
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Resumo do Consumo:', 0, 1)
            
            # Calcular totais
            total_planejado = consumo_mensal['horas_planejadas'].sum()
            total_realizado = consumo_mensal['horas_realizadas'].sum()
            
            # Calcular média mensal e projeção
            meses_passados = consumo_mensal[consumo_mensal['status'] != 'Futuro']
            media_mensal = meses_passados['horas_realizadas'].mean() if not meses_passados.empty else 0
            meses_futuros = len(consumo_mensal[consumo_mensal['status'] == 'Futuro'])
            
            projecao_total = total_realizado + (media_mensal * meses_futuros)
            
            # Exibir métricas resumidas
            col_width = 95
            
            # Linha 1
            pdf.set_font('Arial', 'B', 11)
            pdf.cell(col_width, 8, 'Total Planejado:', 0, 0)
            pdf.cell(col_width, 8, 'Total Realizado:', 0, 1)
            
            pdf.set_font('Arial', '', 11)
            pdf.cell(col_width, 8, format_hours_minutes(total_planejado), 0, 0)
            pdf.cell(col_width, 8, format_hours_minutes(total_realizado), 0, 1)
            
            # Linha 2
            pdf.set_font('Arial', 'B', 11)
            pdf.cell(col_width, 8, 'Percentual Realizado:', 0, 0)
            pdf.cell(col_width, 8, 'Projeção Final:', 0, 1)
            
            pdf.set_font('Arial', '', 11)
            percentual_realizado = (total_realizado / total_planejado * 100) if total_planejado > 0 else 0
            percentual_projecao = (projecao_total / projeto_info['total_hours'] * 100) if projeto_info['total_hours'] > 0 else 0
            
            pdf.cell(col_width, 8, f"{percentual_realizado:.1f}% do planejado", 0, 0)
            pdf.cell(col_width, 8, f"{format_hours_minutes(projecao_total)} ({percentual_projecao:.1f}% do total)", 0, 1)
            
            pdf.ln(3)
            
            # Detalhamento mensal em tabela
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Detalhamento Mensal:', 0, 1)
            
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 8)
            
            # Definir larguras das colunas
            col_widths = [25, 25, 35, 35, 30, 40]
            
            pdf.cell(col_widths[0], 6, 'Mês', 1, 0, 'C', True)
            pdf.cell(col_widths[1], 6, 'Status', 1, 0, 'C', True)
            pdf.cell(col_widths[2], 6, 'Planejado', 1, 0, 'C', True)
            pdf.cell(col_widths[3], 6, 'Realizado', 1, 0, 'C', True)
            pdf.cell(col_widths[4], 6, 'Regulares', 1, 0, 'C', True)
            pdf.cell(col_widths[5], 6, 'Percentual', 1, 1, 'C', True)
            
            # Conteúdo da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 8)
            
            row_count = 0
            for _, row in consumo_mensal.iterrows():
                # Cores baseadas no status
                if row['status'] == 'Atual':
                    pdf.set_fill_color(227, 242, 253)  # Azul claro para mês atual
                    fill = True
                elif row['status'] == 'Passado':
                    pdf.set_fill_color(241, 248, 233)  # Verde claro para meses passados
                    fill = True
                else:
                    if row_count % 2 == 0:
                        pdf.set_fill_color(255, 255, 255)  # Branco
                        fill = False
                    else:
                        pdf.set_fill_color(240, 240, 240)  # Cinza claro
                        fill = True
                
                # Mês
                pdf.cell(col_widths[0], 7, row['mes_str'], 1, 0, 'C', fill)
                
                # Status
                pdf.cell(col_widths[1], 7, row['status'], 1, 0, 'C', fill)
                
                # Horas planejadas
                pdf.cell(col_widths[2], 7, format_hours_minutes(row['horas_planejadas']), 1, 0, 'C', fill)
                
                # Horas realizadas
                pdf.cell(col_widths[3], 7, format_hours_minutes(row['horas_realizadas']), 1, 0, 'C', fill)
                
                # Horas regulares
                pdf.cell(col_widths[4], 7, format_hours_minutes(row['horas_regulares']), 1, 0, 'C', fill)
                
                # Percentual
                pdf.cell(col_widths[5], 7, f"{row['percentual']:.1f}%", 1, 1, 'C', fill)
                
                row_count += 1
            
            # Criar gráfico de consumo mensal
            # Criar gráfico de consumo mensal
            try:
                plt.figure(figsize=(8, 4))
                
                # Preparar dados
                meses = consumo_mensal['mes_str'].tolist()
                planejado = consumo_mensal['horas_planejadas'].tolist()
                realizado = consumo_mensal['horas_realizadas'].tolist()
                
                # Criar gráfico de barras agrupadas
                bar_width = 0.35
                indices = range(len(meses))
                
                plt.bar([i - bar_width/2 for i in indices], planejado, bar_width, label='Planejado', color='#4CAF50')
                plt.bar([i + bar_width/2 for i in indices], realizado, bar_width, label='Realizado', color='#2196F3')
                
                # Diminuir tamanho das fontes
                plt.xlabel('Mês', fontsize=8)
                plt.ylabel('Horas', fontsize=8)
                plt.title('Consumo de Horas Mensais', fontsize=10)
                
                # Diminuir tamanho dos rótulos de eixos
                plt.xticks(indices, meses, rotation=45, ha='right', fontsize=7)
                plt.yticks(fontsize=7)
                
                # Diminuir tamanho da legenda
                plt.legend(fontsize=8)
                
                # Ajustar os tamanhos dos ticks
                plt.tick_params(axis='both', which='major', labelsize=7)
                plt.tick_params(axis='both', which='minor', labelsize=6)
                
                plt.tight_layout()
                
                # Salvar o gráfico em um arquivo temporário
                chart_path = os.path.join(temp_dir, 'consumo_mensal.png')
                plt.savefig(chart_path, dpi=100)  # Aumentar DPI para melhor qualidade
                plt.close()
                
                # Adicionar o gráfico ao PDF
                pdf.ln(2)
                #pdf.set_font('Arial', 'B', 14)
                #pdf.cell(0, 8, 'Gráfico de Consumo Mensal:', 0, 1)
                pdf.image(chart_path, x=10, y=pdf.get_y(), w=190)
                
                # Ajustar posição Y após o gráfico
                pdf.set_y(pdf.get_y() + 130)  # Ajuste aproximado para altura do gráfico
            except Exception as e:
                pdf.ln(3)
                pdf.set_font('Arial', '', 8)
                pdf.cell(0, 10, f"Não foi possível gerar o gráfico de consumo mensal: {str(e)}", 0, 1)
        
        # 4. Análise por Categorias e Atividades
        if (horas_por_categoria is not None and not horas_por_categoria.empty) or (horas_por_atividade is not None and not horas_por_atividade.empty):
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, '4. Análise por Categorias e Atividades', 0, 1)
            
            # Análise por categorias
            if horas_por_categoria is not None and not horas_por_categoria.empty:
                pdf.ln(5)
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(0, 10, 'Horas por Categoria:', 0, 1)
                
                # Cabeçalho da tabela
                pdf.set_fill_color(31, 119, 180)  # Azul
                pdf.set_text_color(255, 255, 255)  # Branco
                pdf.set_font('Arial', 'B', 10)
                
                # Definir larguras das colunas
                col_widths = [80, 30, 40, 40]
                
                pdf.cell(col_widths[0], 8, 'Categoria', 1, 0, 'C', True)
                pdf.cell(col_widths[1], 8, 'Horas', 1, 0, 'C', True)
                pdf.cell(col_widths[2], 8, '% Horas', 1, 0, 'C', True)
                pdf.cell(col_widths[3], 8, 'Custo', 1, 1, 'C', True)
                
                # Conteúdo da tabela
                pdf.set_text_color(0, 0, 0)  # Preto
                pdf.set_font('Arial', '', 9)
                
                # Ordenar por horas (decrescente)
                categorias_ordenadas = horas_por_categoria.sort_values('hours', ascending=False)
                
                row_count = 0
                for _, row in categorias_ordenadas.iterrows():
                    # Alternar cores das linhas
                    if row_count % 2 == 0:
                        pdf.set_fill_color(255, 255, 255)  # Branco
                        fill = False
                    else:
                        pdf.set_fill_color(240, 240, 240)  # Cinza claro
                        fill = True
                    
                    # Categoria (substituir valores nulos por "Sem categoria")
                    categoria = str(row['task_category']) if pd.notna(row['task_category']) else "Sem categoria"
                    pdf.cell(col_widths[0], 7, categoria, 1, 0, 'L', fill)
                    
                    # Horas em formato HH:MM
                    horas_fmt = format_hours_minutes(row['hours'])
                    pdf.cell(col_widths[1], 7, horas_fmt, 1, 0, 'C', fill)
                    
                    # Percentual de horas
                    percentual_horas = row['percentual_horas'] if 'percentual_horas' in row else 0
                    pdf.cell(col_widths[2], 7, f"{percentual_horas:.1f}%", 1, 0, 'C', fill)
                    
                    # Custo
                    custo = row['custo_calculado'] if 'custo_calculado' in row else 0
                    pdf.cell(col_widths[3], 7, f"EUR {custo:.2f}", 1, 1, 'R', fill)
                    
                    row_count += 1
                
                # Gráfico de categorias
                try:
                    plt.figure(figsize=(10, 6))
                    plt.bar(
                        categorias_ordenadas['task_category'].apply(lambda x: str(x) if pd.notna(x) else "Sem categoria"), 
                        categorias_ordenadas['hours'],
                        color='#1E88E5'
                    )
                    plt.title('Distribuição de Horas por Categoria')
                    plt.xlabel('Categoria')
                    plt.ylabel('Horas Trabalhadas')
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    
                    # Salvar o gráfico em um arquivo temporário
                    chart_path = os.path.join(temp_dir, 'categorias_horas.png')
                    plt.savefig(chart_path)
                    plt.close()
                    
                    # Adicionar o gráfico ao PDF
                    pdf.ln(10)
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 10, 'Gráfico de Horas por Categoria:', 0, 1)
                    pdf.image(chart_path, x=10, y=pdf.get_y(), w=190)
                    
                    # Ajustar posição Y após o gráfico
                    pdf.set_y(pdf.get_y() + 130)  # Ajuste aproximado para altura do gráfico
                except Exception as e:
                    pdf.ln(5)
                    pdf.set_font('Arial', '', 10)
                    pdf.cell(0, 10, f"Não foi possível gerar o gráfico de categorias: {str(e)}", 0, 1)
            
            # Análise por atividades
            if horas_por_atividade is not None and not horas_por_atividade.empty:
                pdf.add_page()
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(0, 10, 'Horas por Atividade:', 0, 1)
                
                # Cabeçalho da tabela
                pdf.set_fill_color(31, 119, 180)  # Azul
                pdf.set_text_color(255, 255, 255)  # Branco
                pdf.set_font('Arial', 'B', 10)
                
                # Definir larguras das colunas
                col_widths = [80, 30, 40, 40]
                
                pdf.cell(col_widths[0], 8, 'Atividade', 1, 0, 'C', True)
                pdf.cell(col_widths[1], 8, 'Horas', 1, 0, 'C', True)
                pdf.cell(col_widths[2], 8, '% Horas', 1, 0, 'C', True)
                pdf.cell(col_widths[3], 8, 'Custo', 1, 1, 'C', True)
                
                # Conteúdo da tabela
                pdf.set_text_color(0, 0, 0)  # Preto
                pdf.set_font('Arial', '', 9)
                
                # Ordenar por horas (decrescente)
                atividades_ordenadas = horas_por_atividade.sort_values('hours', ascending=False)
                
                row_count = 0
                for _, row in atividades_ordenadas.iterrows():
                    # Alternar cores das linhas
                    if row_count % 2 == 0:
                        pdf.set_fill_color(255, 255, 255)  # Branco
                        fill = False
                    else:
                        pdf.set_fill_color(240, 240, 240)  # Cinza claro
                        fill = True
                    
                    # Atividade (substituir valores nulos por "Sem atividade")
                    atividade = str(row['activity_name']) if pd.notna(row['activity_name']) else "Sem atividade"
                    pdf.cell(col_widths[0], 7, atividade, 1, 0, 'L', fill)
                    
                    # Horas em formato HH:MM
                    horas_fmt = format_hours_minutes(row['hours'])
                    pdf.cell(col_widths[1], 7, horas_fmt, 1, 0, 'C', fill)
                    
                    # Percentual de horas
                    percentual_horas = row['percentual_horas'] if 'percentual_horas' in row else 0
                    pdf.cell(col_widths[2], 7, f"{percentual_horas:.1f}%", 1, 0, 'C', fill)
                    
                    # Custo
                    custo = row['custo_calculado'] if 'custo_calculado' in row else 0
                    pdf.cell(col_widths[3], 7, f"EUR {custo:.2f}", 1, 1, 'R', fill)
                    
                    row_count += 1
                
                # Gráfico de atividades
                try:
                    plt.figure(figsize=(10, 6))
                    plt.bar(
                        atividades_ordenadas['activity_name'].apply(lambda x: str(x) if pd.notna(x) else "Sem atividade"), 
                        atividades_ordenadas['hours'],
                        color='#43A047'  # Verde para diferenciar do gráfico de categorias
                    )
                    plt.title('Distribuição de Horas por Atividade')
                    plt.xlabel('Atividade')
                    plt.ylabel('Horas Trabalhadas')
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    
                    # Salvar o gráfico em um arquivo temporário
                    chart_path = os.path.join(temp_dir, 'atividades_horas.png')
                    plt.savefig(chart_path)
                    plt.close()
                    
                    # Adicionar o gráfico ao PDF
                    pdf.ln(10)
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 10, 'Gráfico de Horas por Atividade:', 0, 1)
                    pdf.image(chart_path, x=10, y=pdf.get_y(), w=190)
                    
                    # Ajustar posição Y após o gráfico
                    pdf.set_y(pdf.get_y() + 130)  # Ajuste aproximado para altura do gráfico
                except Exception as e:
                    pdf.ln(5)
                    pdf.set_font('Arial', '', 10)
                    pdf.cell(0, 10, f"Não foi possível gerar o gráfico de atividades: {str(e)}", 0, 1)
        
        # 5. Fases do Projeto
        if phases_info is not None and phases_info.get('total_phases', 0) > 0:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, '5. Fases do Projeto', 0, 1)
            
            pdf.ln(5)
            pdf.set_font('Arial', '', 11)
            pdf.cell(0, 10, f"Este projeto possui {phases_info['total_phases']} fases definidas.", 0, 1)
            
            # Tabela de fases
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Detalhamento das Fases:', 0, 1)
            
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 10)
            
            # Definir larguras das colunas
            col_widths = [60, 30, 30, 30, 40]
            
            pdf.cell(col_widths[0], 8, 'Nome', 1, 0, 'C', True)
            pdf.cell(col_widths[1], 8, 'Início', 1, 0, 'C', True)
            pdf.cell(col_widths[2], 8, 'Término', 1, 0, 'C', True)
            pdf.cell(col_widths[3], 8, 'Horas', 1, 0, 'C', True)
            pdf.cell(col_widths[4], 8, 'Status', 1, 1, 'C', True)
            
            # Conteúdo da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 9)
            
            row_count = 0
            for phase in phases_info['phases']:
                # Alternar cores das linhas
                if row_count % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)  # Branco
                    fill = False
                else:
                    pdf.set_fill_color(240, 240, 240)  # Cinza claro
                    fill = True
                
                # Destacar fase atual
                if phase['status'] == 'active':
                    pdf.set_fill_color(227, 242, 253)  # Azul claro
                    fill = True
                
                # Nome da fase
                pdf.cell(col_widths[0], 7, phase['name'], 1, 0, 'L', fill)
                
                # Data de início
                inicio = phase['start_date'].strftime('%d/%m/%Y')
                pdf.cell(col_widths[1], 7, inicio, 1, 0, 'C', fill)
                
                # Data de término
                termino = phase['end_date'].strftime('%d/%m/%Y')
                pdf.cell(col_widths[2], 7, termino, 1, 0, 'C', fill)
                
                # Horas planejadas
                horas_fmt = format_hours_minutes(phase['total_hours'])
                pdf.cell(col_widths[3], 7, horas_fmt, 1, 0, 'C', fill)
                
                # Status
                pdf.cell(col_widths[4], 7, phase['status'].capitalize(), 1, 1, 'C', fill)
                
                row_count += 1
            
            # Análise de progresso das fases
            pdf.ln(10)
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Progresso das Fases:', 0, 1)
            
            # Determinar status para cada fase
            today = datetime.now().date()
            progress_data = []
            
            for phase in phases_info['phases']:
                phase_start = phase['start_date'].date()
                phase_end = phase['end_date'].date()
                phase_duration = (phase_end - phase_start).days
                
                # Calcular progresso esperado
                if today < phase_start:
                    expected_progress = 0  # Fase não iniciada
                elif today > phase_end:
                    expected_progress = 100  # Fase concluída
                else:
                    days_passed = (today - phase_start).days
                    expected_progress = (days_passed / phase_duration * 100) if phase_duration > 0 else 0
                
                # Determinar status da fase
                if phase['status'] == 'completed':
                    actual_progress = 100
                    status = "Concluída"
                elif phase['status'] == 'active':
                    # Para fases ativas, o progresso é proporcional ao tempo
                    actual_progress = expected_progress
                    status = "Em Andamento"
                else:
                    actual_progress = 0
                    status = "Pendente"
                
                # Determinar se está no prazo
                on_schedule = "No Prazo" if actual_progress >= expected_progress else "Atrasada"
                
                progress_data.append({
                    'name': phase['name'],
                    'expected_progress': expected_progress,
                    'actual_progress': actual_progress,
                    'status': status,
                    'on_schedule': on_schedule
                })
            
            # Tabela de progresso
            # Cabeçalho da tabela
            pdf.set_fill_color(31, 119, 180)  # Azul
            pdf.set_text_color(255, 255, 255)  # Branco
            pdf.set_font('Arial', 'B', 10)
            
            # Definir larguras das colunas
            col_widths = [60, 40, 40, 50]
            
            pdf.cell(col_widths[0], 8, 'Fase', 1, 0, 'C', True)
            pdf.cell(col_widths[1], 8, 'Progresso Esperado', 1, 0, 'C', True)
            pdf.cell(col_widths[2], 8, 'Progresso Atual', 1, 0, 'C', True)
            pdf.cell(col_widths[3], 8, 'Situação', 1, 1, 'C', True)
            
            # Conteúdo da tabela
            pdf.set_text_color(0, 0, 0)  # Preto
            pdf.set_font('Arial', '', 9)
            
            for i, data in enumerate(progress_data):
                # Alternar cores das linhas
                if i % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)  # Branco
                    fill = False
                else:
                    pdf.set_fill_color(240, 240, 240)  # Cinza claro
                    fill = True
                
                # Nome da fase
                pdf.cell(col_widths[0], 7, data['name'], 1, 0, 'L', fill)
                
                # Progresso esperado
                pdf.cell(col_widths[1], 7, f"{data['expected_progress']:.1f}%", 1, 0, 'C', fill)
                
                # Progresso atual
                pdf.cell(col_widths[2], 7, f"{data['actual_progress']:.1f}%", 1, 0, 'C', fill)
                
                # Situação (mudar cor do texto dependendo se está no prazo)
                if data['on_schedule'] == "Atrasada":
                    pdf.set_text_color(255, 0, 0)  # Vermelho para atrasado
                else:
                    pdf.set_text_color(0, 128, 0)  # Verde para no prazo
                    
                pdf.cell(col_widths[3], 7, data['on_schedule'], 1, 1, 'C', fill)
                
                # Resetar cor do texto
                pdf.set_text_color(0, 0, 0)
            
            # Criar gráfico de progresso
            try:
                plt.figure(figsize=(10, 6))
                
                # Preparar dados
                phases = [data['name'] for data in progress_data]
                exp_progress = [data['expected_progress'] for data in progress_data]
                act_progress = [data['actual_progress'] for data in progress_data]
                
                # Criar gráfico de barras agrupadas
                bar_width = 0.35
                indices = range(len(phases))
                
                plt.bar([i - bar_width/2 for i in indices], exp_progress, bar_width, label='Esperado', color='#9E9E9E')
                plt.bar([i + bar_width/2 for i in indices], act_progress, bar_width, label='Atual', color='#2196F3')
                
                plt.xlabel('Fase')
                plt.ylabel('Progresso (%)')
                plt.title('Progresso das Fases do Projeto')
                plt.xticks(indices, phases, rotation=45, ha='right')
                plt.ylim(0, 110)  # Deixar um espaço acima de 100%
                plt.legend()
                plt.tight_layout()
                
                # Salvar o gráfico em um arquivo temporário
                chart_path = os.path.join(temp_dir, 'progresso_fases.png')
                plt.savefig(chart_path)
                plt.close()
                
                # Adicionar o gráfico ao PDF
                pdf.ln(10)
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(0, 10, 'Gráfico de Progresso das Fases:', 0, 1)
                pdf.image(chart_path, x=10, y=pdf.get_y(), w=190)
                
            except Exception as e:
                pdf.ln(5)
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 10, f"Não foi possível gerar o gráfico de progresso: {str(e)}", 0, 1)
        
        # Gerar o PDF em memória
        # Gerar o PDF em memória - CORREÇÃO
        output = io.BytesIO()
        pdf.output(dest='S').encode('latin-1')  # Gerar como string e codificar
        output.write(pdf.output(dest='S').encode('latin-1'))  # Escrever bytes codificados no BytesIO
        output.seek(0)  # Voltar ao início do buffer

        # Retornar bytes para download
        return output.getvalue()
    
    except Exception as e:
        st.error(f"Erro ao gerar PDF: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None

def download_project_report(project_info, client_name, db_manager):
    """
    Função para ser chamada na interface do Streamlit para gerar e baixar
    o relatório do projeto em PDF.
    
    Args:
        project_info: Informações do projeto
        client_name: Nome do cliente
        db_manager: Instância do DatabaseManager
    """
    try:
        # Buscar dados necessários para o relatório
        project_id = project_info['project_id']
        
        # Carregar entradas de timesheet
        entries = db_manager.query_to_df(f"SELECT * FROM timesheet WHERE project_id = {project_id}")
        
        # Carregar tabelas complementares
        users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
        rates_df = db_manager.query_to_df("SELECT * FROM rates")
        categories_df = db_manager.query_to_df("SELECT * FROM task_categories")
        activities_df = db_manager.query_to_df("SELECT * FROM activities")
        
        # Verificar e tratar horas migradas (horas_realizadas_mig)
        horas_migradas = 0
        if 'horas_realizadas_mig' in project_info and not pd.isna(project_info['horas_realizadas_mig']):
            horas_migradas = float(project_info['horas_realizadas_mig'])
        
        # Verificar e tratar custo migrado (custo_realizado_mig)
        custo_migrado = 0
        if 'custo_realizado_mig' in project_info and not pd.isna(project_info['custo_realizado_mig']):
            custo_migrado = float(project_info['custo_realizado_mig'])
        
        # Calcular horas regulares e extras
        horas_regulares = 0
        horas_extras = 0
        
        if not entries.empty:
            # Separar horas extras e regulares
            horas_regulares = entries[~entries['overtime'].astype(bool)]['hours'].sum()
            # Para horas extras, vamos guardar o valor original para informação, mas contabilizar o dobro
            horas_extras_originais = entries[entries['overtime'].astype(bool)]['hours'].sum()
            # Calculamos o dobro para contabilização
            horas_extras = horas_extras_originais * 2
        
        # Total de horas (regulares + extras*2)
        horas_realizadas = float(horas_regulares + horas_extras)
        
        # Adicionar horas migradas ao total
        horas_realizadas_total = horas_realizadas + horas_migradas
        
        # Calcular custo realizado
        custo_realizado = 0
        if not entries.empty:
            for _, entry in entries.iterrows():
                try:
                    user_id = entry['user_id']
                    hours = float(entry['hours'])
                    is_overtime = entry['overtime'] if 'overtime' in entry else False
                    
                    # Converter para booleano caso seja outro tipo de dado
                    if isinstance(is_overtime, (int, float)):
                        is_overtime = bool(is_overtime)
                    elif isinstance(is_overtime, str):
                        is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                    
                    # Obter rate para o usuário
                    rate_value = None
                    if 'rate_value' in entry and not pd.isna(entry['rate_value']):
                        rate_value = float(entry['rate_value'])
                    else:
                        user_info = db_manager.query_to_df(f"SELECT rate_id FROM utilizadores WHERE user_id = {user_id}")
                        
                        if not user_info.empty and not pd.isna(user_info['rate_id'].iloc[0]):
                            rate_id = user_info['rate_id'].iloc[0]
                            rate_info = db_manager.query_to_df(f"SELECT rate_cost FROM rates WHERE rate_id = {rate_id}")
                            
                            if not rate_info.empty:
                                rate_value = float(rate_info['rate_cost'].iloc[0])
                    
                    # Calcular custo com base no rate obtido
                    if rate_value:
                        entry_cost = hours * rate_value
                        
                        # Se for hora extra, multiplicar por 2
                        if is_overtime:
                            custo_realizado += entry_cost * 2  # Dobro para horas extras
                        else:
                            custo_realizado += entry_cost  # Normal para horas regulares
                except Exception as e:
                    pass
        
        # Adicionar custo migrado ao total
        custo_realizado_total = custo_realizado + custo_migrado
        
        # Calcular métricas baseadas em dias úteis
        data_inicio = pd.to_datetime(project_info['start_date']).date()
        data_fim = pd.to_datetime(project_info['end_date']).date()
        data_atual = min(datetime.now().date(), data_fim)
        
        dias_uteis_totais = calcular_dias_uteis_projeto(data_inicio, data_fim)
        dias_uteis_decorridos = calcular_dias_uteis_projeto(data_inicio, data_atual)
        dias_uteis_restantes = dias_uteis_totais - dias_uteis_decorridos
        
        percentual_tempo_decorrido = dias_uteis_decorridos / dias_uteis_totais if dias_uteis_totais > 0 else 0
        horas_diarias_planejadas = project_info['total_hours'] / dias_uteis_totais if dias_uteis_totais > 0 else 0
        horas_planejadas_ate_agora = horas_diarias_planejadas * dias_uteis_decorridos
        custo_planejado_ate_agora = project_info['total_cost'] * percentual_tempo_decorrido
        
        # Usar valores totais (incluindo migrados) para CPI
        cpi = custo_planejado_ate_agora / custo_realizado_total if custo_realizado_total > 0 else 1.0
        
        # Usar valores totais para EAC e VAC
        eac = custo_realizado_total
        if dias_uteis_decorridos > 0 and cpi != 0:
            custo_diario_real = custo_realizado_total / dias_uteis_decorridos
            custo_projetado_restante = (custo_diario_real / cpi) * dias_uteis_restantes
            eac = custo_realizado_total + custo_projetado_restante
        vac = project_info['total_cost'] - eac
        
        # Métricas do projeto
        metricas = {
            'cpi': cpi,
            'custo_planejado': custo_planejado_ate_agora,
            'custo_realizado': custo_realizado_total,
            'custo_realizado_atual': custo_realizado,
            'custo_realizado_migrado': custo_migrado,
            'eac': eac,
            'vac': vac,
            'horas_realizadas': horas_realizadas_total,
            'horas_realizadas_atual': horas_realizadas,
            'horas_realizadas_migrado': horas_migradas,
            'horas_regulares': horas_regulares,
            'horas_extras_originais': horas_extras_originais if 'horas_extras_originais' in locals() else 0,
            'horas_extras': horas_extras,
            'dias_uteis_totais': dias_uteis_totais,
            'dias_uteis_decorridos': dias_uteis_decorridos,
            'dias_uteis_restantes': dias_uteis_restantes,
            'horas_diarias_planejadas': horas_diarias_planejadas,
            'horas_planejadas_ate_agora': horas_planejadas_ate_agora
        }
        
        # Preparar informações de recursos
        recursos_dados = []
        
        # Adicionar dados do sistema atual
        if not entries.empty:
            # Mesclar com informações de usuários para obter nomes
            recursos_df = entries.merge(
                users_df[['user_id', 'First_Name', 'Last_Name', 'rate_id']],
                on='user_id',
                how='left'
            )
            
            # Adicionar informações de custo baseadas nas rates, considerando horas extras
            def calcular_custo_entrada(row):
                if pd.isna(row['rate_id']):
                    return 0
                
                rate_info = rates_df[rates_df['rate_id'] == row['rate_id']]
                if rate_info.empty:
                    return 0
                
                rate_value = float(rate_info['rate_cost'].iloc[0])
                is_overtime = row.get('overtime', False)
                
                # Converter is_overtime para booleano
                if isinstance(is_overtime, (int, float)):
                    is_overtime = bool(is_overtime)
                elif isinstance(is_overtime, str):
                    is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                
                # Multiplicar por 2 se for hora extra
                multiplicador = 2 if is_overtime else 1
                return float(row['hours']) * rate_value * multiplicador
            
            recursos_df['custo'] = recursos_df.apply(calcular_custo_entrada, axis=1)
            
            # Coluna para indicar se é hora extra
            recursos_df['is_extra'] = recursos_df.apply(
                lambda row: bool(row.get('overtime', False)), 
                axis=1
            )
            
            # Criar coluna de nome completo
            recursos_df['nome_completo'] = recursos_df.apply(
                lambda row: f"{row['First_Name']} {row['Last_Name']}" if pd.notna(row['First_Name']) else "Desconhecido",
                axis=1
            )
            
            # Agrupar por recurso (usuário)
            for nome, dados in recursos_df.groupby('nome_completo'):
                horas_total = dados['hours'].sum()
                custo_total = dados['custo'].sum()
                recursos_dados.append({
                    'nome_completo': nome,
                    'hours': horas_total,
                    'custo': custo_total
                })
        
        # Adicionar dados migrados como recurso "GLPI" se existirem
        if metricas['horas_realizadas_migrado'] > 0 or metricas['custo_realizado_migrado'] > 0:
            recursos_dados.append({
                'nome_completo': 'GLPI',
                'hours': metricas['horas_realizadas_migrado'],
                'custo': metricas['custo_realizado_migrado']
            })
        
        # Converter para DataFrame
        recursos_utilizacao = pd.DataFrame(recursos_dados)
        
        if not recursos_utilizacao.empty:
            # Adicionar percentual sobre o total
            total_horas = recursos_utilizacao['hours'].sum()
            total_custo = recursos_utilizacao['custo'].sum()
            
            recursos_utilizacao['perc_horas'] = recursos_utilizacao['hours'].apply(
                lambda x: (x / total_horas * 100) if total_horas > 0 else 0
            )
            recursos_utilizacao['perc_custo'] = recursos_utilizacao['custo'].apply(
                lambda x: (x / total_custo * 100) if total_custo > 0 else 0
            )
        
        # Preparar dados de categorias e atividades
        horas_por_categoria = None
        horas_por_atividade = None
        
        if not entries.empty:
            timesheet_completo = entries.copy()
            
            # Juntar com categorias (se category_id existir nas entradas)
            if 'category_id' in timesheet_completo.columns:
                timesheet_completo = timesheet_completo.merge(
                    categories_df[['task_category_id', 'task_category']],
                    left_on='category_id',
                    right_on='task_category_id',
                    how='left'
                )
            
            # Juntar com atividades (se task_id existir nas entradas)
            if 'task_id' in timesheet_completo.columns:
                timesheet_completo = timesheet_completo.merge(
                    activities_df[['activity_id', 'activity_name']],
                    left_on='task_id',
                    right_on='activity_id',
                    how='left'
                )
            
            # Juntar com informações de usuários para obter rates
            timesheet_completo = timesheet_completo.merge(
                users_df[['user_id', 'First_Name', 'Last_Name', 'rate_id']],
                on='user_id',
                how='left'
            )
            
            # Função para calcular custo baseado na rate do usuário e horas extras
            def calcular_custo(row):
                if pd.isna(row['rate_id']):
                    return 0
                
                rate_info = rates_df[rates_df['rate_id'] == row['rate_id']]
                if rate_info.empty:
                    return 0
                
                rate_value = float(rate_info['rate_cost'].iloc[0])
                is_overtime = row.get('overtime', False)
                
                # Converter is_overtime para booleano
                if isinstance(is_overtime, (int, float)):
                    is_overtime = bool(is_overtime)
                elif isinstance(is_overtime, str):
                    is_overtime = is_overtime.lower() in ('true', 't', 'yes', 'y', '1')
                
                # Multiplicar por 2 se for hora extra
                multiplicador = 2 if is_overtime else 1
                return float(row['hours']) * rate_value * multiplicador
            
            timesheet_completo['custo_calculado'] = timesheet_completo.apply(calcular_custo, axis=1)
            
            # Verificar se há informações de categoria
            if 'task_category' in timesheet_completo.columns:
                # Agrupar por categoria
                horas_por_categoria = timesheet_completo.groupby('task_category').agg({
                    'hours': 'sum',
                    'custo_calculado': 'sum'
                }).reset_index()
                
                # Calcular percentuais
                total_horas_cat = horas_por_categoria['hours'].sum()
                total_custo_cat = horas_por_categoria['custo_calculado'].sum()
                
                horas_por_categoria['percentual_horas'] = horas_por_categoria['hours'].apply(
                    lambda x: (x / total_horas_cat * 100) if total_horas_cat > 0 else 0
                )
                
                horas_por_categoria['percentual_custo'] = horas_por_categoria['custo_calculado'].apply(
                    lambda x: (x / total_custo_cat * 100) if total_custo_cat > 0 else 0
                )
                
                # Substituir valores NaN na categoria
                horas_por_categoria['task_category'] = horas_por_categoria['task_category'].fillna('Sem categoria')
            
            # Verificar se há informações de atividade
            if 'activity_name' in timesheet_completo.columns:
                # Agrupar por atividade
                horas_por_atividade = timesheet_completo.groupby('activity_name').agg({
                    'hours': 'sum',
                    'custo_calculado': 'sum'
                }).reset_index()
                
                # Calcular percentuais
                total_horas_act = horas_por_atividade['hours'].sum()
                total_custo_act = horas_por_atividade['custo_calculado'].sum()
                
                horas_por_atividade['percentual_horas'] = horas_por_atividade['hours'].apply(
                    lambda x: (x / total_horas_act * 100) if total_horas_act > 0 else 0
                )
                
                horas_por_atividade['percentual_custo'] = horas_por_atividade['custo_calculado'].apply(
                    lambda x: (x / total_custo_act * 100) if total_custo_act > 0 else 0
                )
                
                # Substituir valores NaN na atividade
                horas_por_atividade['activity_name'] = horas_por_atividade['activity_name'].fillna('Sem atividade')
        
        # Calcular dados de consumo mensal
        consumo_mensal = None
        try:
            # Período do projeto
            data_inicio_projeto = pd.to_datetime(project_info['start_date'], format='mixed')
            data_fim_projeto = pd.to_datetime(project_info['end_date'], format='mixed')
            
            # Criar uma lista de meses entre o início e fim do projeto
            meses = []
            atual = data_inicio_projeto.replace(day=1)
            while atual <= data_fim_projeto:
                meses.append(atual)
                # Avançar para o próximo mês
                if atual.month == 12:
                    atual = atual.replace(year=atual.year + 1, month=1)
                else:
                    atual = atual.replace(month=atual.month + 1)
            
            # Calcular horas planejadas por mês (distribuição uniforme)
            total_meses = len(meses)
            horas_por_mes_planejadas = project_info['total_hours'] / total_meses if total_meses > 0 else 0
            
            # Preparar dados para o gráfico de consumo mensal
            consumo_mensal = []
            
            for mes in meses:
                # Determinar o último dia do mês
                ultimo_dia_mes = calendar.monthrange(mes.year, mes.month)[1]
                fim_mes = mes.replace(day=ultimo_dia_mes)
                
                # Calcular dias úteis no mês
                dias_uteis_mes = calcular_dias_uteis_projeto(mes.date(), fim_mes.date())
                
                # Horas planejadas para o mês (pode ser refinado com base em dias úteis)
                horas_planejadas_mes = (dias_uteis_mes / dias_uteis_totais) * project_info['total_hours'] if dias_uteis_totais > 0 else 0
                
                # Filtrar entradas de timesheet para o mês
                if not entries.empty:
                    # Converter start_date para datetime
                    entries['start_date_dt'] = pd.to_datetime(entries['start_date'], format='mixed')

                    # Use the converted datetime column for filtering
                    entries_mes = entries[
                        (entries['start_date_dt'].dt.year == mes.year) &
                        (entries['start_date_dt'].dt.month == mes.month)
                    ]
                    
                    # Calcular horas registradas no mês
                    horas_regulares_mes = entries_mes[~entries_mes['overtime'].astype(bool)]['hours'].sum() if not entries_mes.empty else 0
                    horas_extras_mes_orig = entries_mes[entries_mes['overtime'].astype(bool)]['hours'].sum() if not entries_mes.empty else 0
                    horas_extras_mes = horas_extras_mes_orig * 2
                    horas_realizadas_mes = horas_regulares_mes + horas_extras_mes
                else:
                    horas_realizadas_mes = 0
                    horas_regulares_mes = 0
                    horas_extras_mes_orig = 0
                    horas_extras_mes = 0
                
                # Tratar meses passados vs. futuros
                data_atual = datetime.now()
                
                # Adicionar horas migradas apenas para o primeiro mês
                horas_migradas_mes = horas_migradas if mes == meses[0] else 0
                
                if mes.year < data_atual.year or (mes.year == data_atual.year and mes.month < data_atual.month):
                    # Mês passado - usar dados reais
                    status = "Passado"
                elif mes.year == data_atual.year and mes.month == data_atual.month:
                    # Mês atual - em andamento
                    status = "Atual"
                else:
                    # Mês futuro - planejado
                    status = "Futuro"
                
                # Calcular percentual em relação ao planejado
                horas_realizadas_total_mes = horas_realizadas_mes + horas_migradas_mes
                percentual = (horas_realizadas_total_mes / horas_planejadas_mes * 100) if horas_planejadas_mes > 0 else 0
                
                # Adicionar ao array de consumo mensal
                consumo_mensal.append({
                    'mes': mes,
                    'mes_str': mes.strftime('%b/%Y'),
                    'horas_planejadas': horas_planejadas_mes,
                    'horas_realizadas': horas_realizadas_total_mes,
                    'horas_regulares': horas_regulares_mes,
                    'horas_extras_orig': horas_extras_mes_orig,
                    'horas_extras': horas_extras_mes,
                    'horas_migradas': horas_migradas_mes,
                    'percentual': percentual,
                    'status': status
                })
            
            # Converter para DataFrame
            consumo_mensal = pd.DataFrame(consumo_mensal)
        except Exception as e:
            st.warning(f"Não foi possível calcular o consumo mensal: {str(e)}")
            consumo_mensal = None
        
        # Obter informações das fases
        try:
            from project_phases import integrate_phases_with_project_reports
            phases_info = integrate_phases_with_project_reports(project_id)
        except Exception as e:
            phases_info = None
        
        # Gerar o PDF
        pdf_bytes = generate_single_project_pdf(
            project_info, 
            metricas, 
            entries, 
            recursos_utilizacao, 
            horas_por_categoria, 
            horas_por_atividade, 
            consumo_mensal, 
            phases_info,
            users_df,
            rates_df,
            client_name
        )
        
        # Nome do arquivo para download
        filename = f"relatorio_{project_info['project_name'].replace(' ', '_')}.pdf"
        
        # Retornar o botão de download
        return st.download_button(
            label="📥 Baixar Relatório do Projeto (PDF)",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf"
        )
        
    except Exception as e:
        st.error(f"Erro ao preparar relatório: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None