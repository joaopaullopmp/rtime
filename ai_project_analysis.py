# ai_project_analysis.py - Versão adaptada para usar banco de dados SQLite
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px
from database_manager import DatabaseManager
from typing import Dict, List, Any, Optional, Union, Tuple

class ProjectAIAnalyzer:
    def __init__(self):
        try:
            # Inicializar o gerenciador de banco de dados
            self.db_manager = DatabaseManager()
            
            # Carregar dados das tabelas necessárias
            self.timesheet_df = self.db_manager.query_to_df("SELECT * FROM timesheet")
            self.users_df = self.db_manager.query_to_df("SELECT * FROM utilizadores")
            self.projects_df = self.db_manager.query_to_df("SELECT * FROM projects")
            self.clients_df = self.db_manager.query_to_df("SELECT * FROM clients")
            self.rates_df = self.db_manager.query_to_df("SELECT * FROM rates")
            
            # Preparar dados
            # No método __init__ da classe ProjectAIAnalyzer, modifique as linhas que fazem a conversão de data:
            if not self.timesheet_df.empty:
                self.timesheet_df['start_date'] = pd.to_datetime(self.timesheet_df['start_date'], format='mixed')
                self.timesheet_df['end_date'] = pd.to_datetime(self.timesheet_df['end_date'], format='mixed')
                
            if not self.projects_df.empty:
                self.projects_df['start_date'] = pd.to_datetime(self.projects_df['start_date'], format='mixed')
                self.projects_df['end_date'] = pd.to_datetime(self.projects_df['end_date'], format='mixed')
                
        except Exception as e:
            # Em caso de erro, inicializar com DataFrames vazios
            self.timesheet_df = pd.DataFrame()
            self.users_df = pd.DataFrame()
            self.projects_df = pd.DataFrame()
            self.clients_df = pd.DataFrame()
            self.rates_df = pd.DataFrame()
            raise e
    
    # Removendo as anotações de tipo para evitar erros
    def calculate_health_score(self, project_entries, metricas):
        """Calcula uma pontuação de saúde do projeto"""
        try:
            if project_entries.empty:
                return 5.0  # Pontuação padrão
            
            # Fatores de saúde
            cpi_score = 10.0 if metricas['cpi'] >= 1.1 else 7.5 if metricas['cpi'] >= 0.95 else 5.0 if metricas['cpi'] >= 0.85 else 2.5
            
            # Retornar pontuação média
            return cpi_score
        except Exception as e:
            print(f"Erro no cálculo de saúde: {str(e)}")
            return 5.0
    
    def analyze_project_health(self, project_id, metrics):
        """Analisa a saúde geral do projeto baseado em métricas e dados históricos"""
        try:
            # Dados do projeto
            project_data = self.projects_df[self.projects_df['project_id'] == project_id].iloc[0] if not self.projects_df.empty else {}
            
            # Dados de timesheet do projeto
            if not self.timesheet_df.empty:
                project_timesheet = self.timesheet_df[self.timesheet_df['project_id'] == project_id]
                
                # Se não houver dados de timesheet, retornar um dicionário vazio
                if project_timesheet.empty:
                    return {}
            else:
                return {}
            
            # Calcular percentual de conclusão baseado nas horas
            total_hours = float(project_data['total_hours']) if 'total_hours' in project_data else 0
            hours_used = metrics['horas_realizadas'] if 'horas_realizadas' in metrics else 0
            completion_percentage = (hours_used / total_hours * 100) if total_hours > 0 else 0
            
            # Determinar saúde do projeto
            schedule_health = "Bom" if metrics['cpi'] >= 1 else "Médio" if metrics['cpi'] >= 0.85 else "Ruim"
            budget_health = "Bom" if metrics['vac'] >= 0 else "Médio" if metrics['vac'] >= -0.1 * float(project_data['total_cost']) else "Ruim"
            
            # Avaliar problemas de recursos
            resource_issues = []
            
            if not self.users_df.empty:
                # Calcular utilização de recursos
                resource_utilization = self._calculate_resource_utilization(project_id)
                
                # Identificar recursos sobreutilizados (> 90%)
                overutilized = [res for res in resource_utilization if res['utilization'] > 90]
                if overutilized:
                    resource_issues.append({
                        'type': 'overutilization',
                        'message': f"{len(overutilized)} recursos estão com utilização acima de 90%",
                        'details': overutilized
                    })
                
                # Identificar recursos subutilizados (< 50%)
                underutilized = [res for res in resource_utilization if res['utilization'] < 50]
                if underutilized:
                    resource_issues.append({
                        'type': 'underutilization',
                        'message': f"{len(underutilized)} recursos estão com utilização abaixo de 50%",
                        'details': underutilized
                    })
            
            # Análise de tendência
            trend_analysis = self._analyze_project_trend(project_id, metrics)
            
            # Dados do cliente
            client_id = project_data['client_id'] if 'client_id' in project_data else None
            client_data = self.clients_df[self.clients_df['client_id'] == client_id].iloc[0] if client_id and not self.clients_df.empty else {}
            
            # Análise de risco
            risk_assessment = self._assess_project_risk(metrics, completion_percentage)
            
            # Coletar tudo em um dicionário de análise
            analysis = {
                'project_name': project_data['project_name'] if 'project_name' in project_data else 'Projeto Desconhecido',
                'client_name': client_data['name'] if 'name' in client_data else 'Cliente Desconhecido',
                'completion': completion_percentage,
                'schedule_health': schedule_health,
                'budget_health': budget_health,
                'resource_issues': resource_issues,
                'trend_analysis': trend_analysis,
                'risk_assessment': risk_assessment,
                'resource_balance': {
                    'utilization_data': resource_utilization if 'resource_utilization' in locals() else []
                }
            }
            
            return analysis
            
        except Exception as e:
            print(f"Erro na análise de saúde do projeto: {str(e)}")
            return {}
    
    def _calculate_resource_utilization(self, project_id):
        """Calcula a utilização de recursos no projeto"""
        try:
            # Filtrar dados de timesheet para o projeto específico
            project_data = self.timesheet_df[self.timesheet_df['project_id'] == project_id]
            
            if project_data.empty or self.users_df.empty:
                return []
            
            # Agrupar por usuário
            user_hours = project_data.groupby('user_id')['hours'].sum().reset_index()
            
            # Juntar com informações do usuário
            user_data = user_hours.merge(
                self.users_df[['user_id', 'First_Name', 'Last_Name', 'rate_id']], 
                on='user_id', 
                how='left'
            )
            
            # Calcular utilização (considerando 8h/dia, 22 dias/mês)
            standard_monthly_hours = 8 * 22  # 176 horas/mês
            
            # Filtrar por período (último mês)
            start_date = datetime.now() - timedelta(days=30)
            recent_data = self.timesheet_df[self.timesheet_df['start_date'] >= start_date]
            
            result = []
            for _, user in user_data.iterrows():
                if pd.isna(user['First_Name']) or pd.isna(user['Last_Name']):
                    continue
                    
                user_name = f"{user['First_Name']} {user['Last_Name']}"
                
                # Horas totais do usuário (em todos os projetos)
                if not recent_data.empty:
                    total_user_hours = recent_data[recent_data['user_id'] == user['user_id']]['hours'].sum()
                    
                    # Horas em projetos faturáveis
                    billable_hours = recent_data[
                        (recent_data['user_id'] == user['user_id']) & 
                        (recent_data['billable'] == True)
                    ]['hours'].sum()
                else:
                    total_user_hours = 0
                    billable_hours = 0
                
                # Calcular utilização
                utilization = (total_user_hours / standard_monthly_hours) * 100
                billable_rate = (billable_hours / total_user_hours) * 100 if total_user_hours > 0 else 0
                
                # Adicionar ao resultado
                result.append({
                    'user_id': int(user['user_id']),
                    'name': user_name,
                    'project_hours': float(user['hours']),
                    'total_hours': float(total_user_hours),
                    'utilization': float(utilization),
                    'billable_rate': float(billable_rate)
                })
            
            return result
            
        except Exception as e:
            print(f"Erro no cálculo de utilização de recursos: {str(e)}")
            return []
    
    def _analyze_project_trend(self, project_id, metrics):
        """Analisa a tendência do projeto com base nos dados históricos"""
        try:
            # Se não houver dados suficientes, retornar resultados vazios
            if self.timesheet_df.empty:
                return {
                    'hours_trend': 'Estável',
                    'cost_trend': 'Estável',
                    'performance_trend': 'Estável'
                }
            
            # Filtrar dados do projeto
            project_data = self.timesheet_df[self.timesheet_df['project_id'] == project_id]
            
            if project_data.empty:
                return {
                    'hours_trend': 'Estável',
                    'cost_trend': 'Estável',
                    'performance_trend': 'Estável'
                }
            
            # Calcular tendências
            # Agrupamento por semana para calcular a tendência
            project_data['week'] = project_data['start_date'].dt.isocalendar().week
            project_data['year'] = project_data['start_date'].dt.isocalendar().year
            
            # Criar chave ano-semana para ordenar corretamente
            project_data['year_week'] = project_data['year'].astype(str) + '-' + project_data['week'].astype(str)
            
            # Agrupar por semana
            weekly_data = project_data.groupby('year_week').agg({
                'hours': 'sum'
            }).reset_index()
            
            # Ordenar por ano-semana
            weekly_data = weekly_data.sort_values('year_week')
            
            # Calcular a tendência de horas
            if len(weekly_data) >= 3:
                # Usar as últimas semanas para determinar a tendência
                recent_weeks = weekly_data.tail(3)
                hours_trend = recent_weeks['hours'].diff().mean()
                
                # Interpretar a tendência
                if hours_trend > 0.5:
                    hours_trend_label = 'Crescente'
                elif hours_trend < -0.5:
                    hours_trend_label = 'Decrescente'
                else:
                    hours_trend_label = 'Estável'
            else:
                hours_trend_label = 'Dados insuficientes'
            
            # Tendência de custos (simplificada)
            cost_trend_label = 'Crescente' if metrics['eac'] > float(self.projects_df[self.projects_df['project_id'] == project_id]['total_cost'].iloc[0]) else 'Decrescente' if metrics['eac'] < float(self.projects_df[self.projects_df['project_id'] == project_id]['total_cost'].iloc[0]) else 'Estável'
            
            # Tendência de performance (baseada no CPI)
            performance_trend_label = 'Boa' if metrics['cpi'] >= 1 else 'Regular' if metrics['cpi'] >= 0.85 else 'Ruim'
            
            return {
                'hours_trend': hours_trend_label,
                'cost_trend': cost_trend_label,
                'performance_trend': performance_trend_label
            }
            
        except Exception as e:
            print(f"Erro na análise de tendência do projeto: {str(e)}")
            return {
                'hours_trend': 'Erro na análise',
                'cost_trend': 'Erro na análise',
                'performance_trend': 'Erro na análise'
            }
    
    def _assess_project_risk(self, metrics, completion_percentage):
        """Avalia o nível de risco do projeto"""
        try:
            # Fatores de risco
            risk_factors = []
            
            # 1. CPI baixo indica risco financeiro
            if metrics['cpi'] < 0.85:
                risk_factors.append({
                    'factor': 'CPI Baixo',
                    'description': 'O projeto está gastando mais que o planejado',
                    'severity': 'Alto' if metrics['cpi'] < 0.7 else 'Médio'
                })
            
            # 2. Alto consumo de horas em relação à conclusão
            # Verificar se o % de horas usadas é desproporcional ao % de conclusão (estimado)
            expected_hours_percentage = completion_percentage
            actual_hours_percentage = (metrics['horas_realizadas'] / metrics['horas_planejadas_ate_agora']) * 100 if metrics['horas_planejadas_ate_agora'] > 0 else 0
            
            hours_deviation = actual_hours_percentage - expected_hours_percentage
            if hours_deviation > 15:
                risk_factors.append({
                    'factor': 'Consumo Excessivo de Horas',
                    'description': 'O consumo de horas está acima do esperado para o progresso atual',
                    'severity': 'Alto' if hours_deviation > 25 else 'Médio'
                })
            
            # 3. Risco de cronograma (baseado nos dias restantes vs horas restantes)
            hours_per_day_remaining = metrics['horas_diarias_planejadas']
            hours_remaining = metrics['horas_planejadas_ate_agora'] - metrics['horas_realizadas']
            
            if hours_per_day_remaining > 0 and metrics['dias_uteis_restantes'] > 0:
                required_daily_rate = hours_remaining / metrics['dias_uteis_restantes']
                rate_deviation = (required_daily_rate / hours_per_day_remaining) - 1
                
                if rate_deviation > 0.2:
                    risk_factors.append({
                        'factor': 'Risco de Cronograma',
                        'description': f"Necessário aumento de {rate_deviation*100:.1f}% na taxa diária de trabalho",
                        'severity': 'Alto' if rate_deviation > 0.5 else 'Médio'
                    })
            
            # Determinar o nível geral de risco
            if any(factor['severity'] == 'Alto' for factor in risk_factors):
                risk_level = 'Alto'
            elif len(risk_factors) > 1:
                risk_level = 'Médio'
            elif len(risk_factors) == 1:
                risk_level = 'Baixo'
            else:
                risk_level = 'Muito Baixo'
            
            return {
                'level': risk_level,
                'factors': risk_factors
            }
            
        except Exception as e:
            print(f"Erro na avaliação de risco do projeto: {str(e)}")
            return {
                'level': 'Indeterminado',
                'factors': [{
                    'factor': 'Erro de Análise',
                    'description': 'Não foi possível completar a análise de risco',
                    'severity': 'Desconhecido'
                }]
            }
    
    def generate_recommendations(self, project_id, metrics):
        """Gera recomendações baseadas na análise do projeto"""
        try:
            # Analisar o projeto
            project_analysis = self.analyze_project_health(project_id, metrics)
            
            # Lista de recomendações
            recommendations = []
            
            # Se não há análise, retornar lista vazia
            if not project_analysis:
                return recommendations
            
            # 1. Recomendações baseadas no CPI
            if project_analysis['budget_health'] == 'Ruim':
                recommendations.append({
                    'title': 'Revisão orçamentária urgente',
                    'description': 'O projeto está significativamente sobre o orçamento. Realize uma revisão detalhada dos gastos e ajuste o escopo ou negocie um aditivo com o cliente.',
                    'priority': 'Alta'
                })
            elif project_analysis['budget_health'] == 'Médio':
                recommendations.append({
                    'title': 'Monitoramento orçamentário',
                    'description': 'O projeto está ligeiramente sobre o orçamento. Implemente controles mais rigorosos de aprovação de horas e recursos.',
                    'priority': 'Média'
                })
            
            # 2. Recomendações baseadas em recursos
            if 'resource_issues' in project_analysis:
                for issue in project_analysis['resource_issues']:
                    if issue['type'] == 'overutilization':
                        recommendations.append({
                            'title': 'Balanceamento de equipe',
                            'description': 'Alguns recursos estão sobrecarregados. Considere redistribuir tarefas ou adicionar recursos ao projeto.',
                            'priority': 'Alta' if len(issue['details']) > 2 else 'Média'
                        })
                    elif issue['type'] == 'underutilization':
                        recommendations.append({
                            'title': 'Otimização de recursos',
                            'description': 'Alguns recursos estão subutilizados. Avalie a possibilidade de reduzir a alocação ou substituir por recursos mais adequados.',
                            'priority': 'Média'
                        })
            
            # 3. Recomendações baseadas na análise de risco
            if project_analysis['risk_assessment']['level'] in ['Alto', 'Médio']:
                risk_desc = "Realize uma revisão completa do projeto, incluindo escopo, cronograma e orçamento. Considere uma reunião especial com o cliente para realinhar expectativas."
                recommendations.append({
                    'title': 'Mitigação de riscos',
                    'description': risk_desc,
                    'priority': 'Alta' if project_analysis['risk_assessment']['level'] == 'Alto' else 'Média'
                })
            
            # 4. Recomendações baseadas na tendência de horas
            if 'trend_analysis' in project_analysis and project_analysis['trend_analysis']['hours_trend'] == 'Crescente':
                recommendations.append({
                    'title': 'Verificação de escopo',
                    'description': 'O consumo de horas está aumentando. Verifique se o escopo está crescendo sem controle ("scope creep") ou se há ineficiências no processo.',
                    'priority': 'Média'
                })
            
            # 5. Recomendações baseadas na conclusão vs. tempo restante
            if 'completion' in project_analysis:
                time_percentage = (metrics['dias_uteis_decorridos'] / metrics['dias_uteis_totais']) * 100 if metrics['dias_uteis_totais'] > 0 else 0
                if time_percentage - project_analysis['completion'] > 15:
                    recommendations.append({
                        'title': 'Revisão de cronograma',
                        'description': 'O progresso do projeto está atrasado em relação ao tempo decorrido. Considere ajustar o cronograma ou aumentar temporariamente os recursos.',
                        'priority': 'Alta' if (time_percentage - project_analysis['completion']) > 25 else 'Média'
                    })
            
            # Ordenar por prioridade (Alta -> Média -> Baixa)
            priority_order = {'Alta': 0, 'Média': 1, 'Baixa': 2}
            recommendations.sort(key=lambda x: priority_order[x['priority']])
            
            return recommendations
            
        except Exception as e:
            print(f"Erro ao gerar recomendações: {str(e)}")
            return []

def render_ai_analysis(project_id, metricas):
    """Renderiza a análise de IA do projeto"""
    try:
        analyzer = ProjectAIAnalyzer()
        analysis = analyzer.analyze_project_health(project_id, metricas)
        recommendations = analyzer.generate_recommendations(project_id, metricas)
        
        # Se não houver análise, mostrar aviso
        if not analysis:
            st.warning("Dados insuficientes para análise inteligente.")
            return
        
        # Dashboard de IA
        st.subheader("🧠 Análise Inteligente")
        
        # Cards de status
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Saúde do orçamento
            budget_colors = {
                'Bom': '#4CAF50',
                'Médio': '#FFC107',
                'Ruim': '#F44336'
            }
            budget_color = budget_colors.get(analysis.get('budget_health', 'Médio'), '#FFC107')
            
            st.markdown(
                f"""
                <div style="padding:1rem; border-radius:0.5rem; background-color:{budget_color}20; 
                border-left:0.5rem solid {budget_color};">
                <h4 style="margin:0; color:{budget_color};">Orçamento</h4>
                <p style="font-size:1.5rem; font-weight:bold; margin:0;">{analysis.get('budget_health', 'Médio')}</p>
                </div>
                """, 
                unsafe_allow_html=True
            )
        
        with col2:
            # Saúde do cronograma
            schedule_colors = {
                'Bom': '#4CAF50',
                'Médio': '#FFC107',
                'Ruim': '#F44336'
            }
            schedule_color = schedule_colors.get(analysis.get('schedule_health', 'Médio'), '#FFC107')
            
            st.markdown(
                f"""
                <div style="padding:1rem; border-radius:0.5rem; background-color:{schedule_color}20; 
                border-left:0.5rem solid {schedule_color};">
                <h4 style="margin:0; color:{schedule_color};">Cronograma</h4>
                <p style="font-size:1.5rem; font-weight:bold; margin:0;">{analysis.get('schedule_health', 'Médio')}</p>
                </div>
                """, 
                unsafe_allow_html=True
            )
        
        with col3:
            # Nível de Risco
            risk_level = analysis.get('risk_assessment', {}).get('level', 'Médio')
            risk_colors = {
                'Muito Baixo': '#4CAF50',
                'Baixo': '#8BC34A',
                'Médio': '#FFC107',
                'Alto': '#F44336',
                'Indeterminado': '#9E9E9E'
            }
            risk_color = risk_colors.get(risk_level, '#FFC107')
            
            st.markdown(
                f"""
                <div style="padding:1rem; border-radius:0.5rem; background-color:{risk_color}20; 
                border-left:0.5rem solid {risk_color};">
                <h4 style="margin:0; color:{risk_color};">Risco</h4>
                <p style="font-size:1.5rem; font-weight:bold; margin:0;">{risk_level}</p>
                </div>
                """, 
                unsafe_allow_html=True
            )
        
        # Análise de tendências
        st.markdown("### 📈 Tendências")
        
        trend_analysis = analysis.get('trend_analysis', {})
        trend_col1, trend_col2, trend_col3 = st.columns(3)
        
        with trend_col1:
            hours_trend = trend_analysis.get('hours_trend', 'Estável')
            hours_icons = {
                'Crescente': '🔺',
                'Decrescente': '🔻',
                'Estável': '➡️',
                'Dados insuficientes': '❓',
                'Erro na análise': '⚠️'
            }
            st.metric("Consumo de Horas", hours_trend, delta=hours_icons.get(hours_trend, ''))
        
        with trend_col2:
            cost_trend = trend_analysis.get('cost_trend', 'Estável')
            cost_icons = {
                'Crescente': '🔺',
                'Decrescente': '🔻',
                'Estável': '➡️',
                'Erro na análise': '⚠️'
            }
            st.metric("Tendência de Custo", cost_trend, delta=cost_icons.get(cost_trend, ''))
        
        with trend_col3:
            perf_trend = trend_analysis.get('performance_trend', 'Regular')
            perf_icons = {
                'Boa': '✅',
                'Regular': '⚠️',
                'Ruim': '❌',
                'Erro na análise': '⚠️'
            }
            st.metric("Performance", perf_trend, delta=perf_icons.get(perf_trend, ''))
            
        # Recursos e utilização
        st.markdown("### 👥 Utilização de Recursos")
        
        resource_balance = analysis.get('resource_balance', {})
        utilization_data = resource_balance.get('utilization_data', [])
        
        if utilization_data:
            # Criar dados para o gráfico
            data = pd.DataFrame(utilization_data)
            
            fig = px.bar(
                data, 
                x='name', 
                y=['utilization', 'billable_rate'],
                title='Ocupação por Recurso',
                labels={'value': 'Percentual (%)', 'name': 'Recursos', 'variable': 'Métrica'},
                barmode='group',
                color_discrete_map={'utilization': '#1E88E5', 'billable_rate': '#4CAF50'}
            )
            
            # Adicionar linhas de referência
            fig.add_shape(
                type='line',
                x0=-0.5,
                x1=len(data)-0.5,
                y0=87.5,
                y1=87.5,
                line=dict(color='red', width=2, dash='dash'),
                name='Meta Ocupação'
            )
            
            fig.add_shape(
                type='line',
                x0=-0.5,
                x1=len(data)-0.5,
                y0=75,
                y1=75,
                line=dict(color='orange', width=2, dash='dash'),
                name='Meta Faturável'
            )
            
            # Mostrar o gráfico
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Dados de utilização de recursos não disponíveis.")
        
        # Recomendações
        st.markdown("### 💡 Recomendações")
        
        if recommendations:
            for i, rec in enumerate(recommendations):
                color = '#F44336' if rec['priority'] == 'Alta' else '#FFC107' if rec['priority'] == 'Média' else '#4CAF50'
                
                st.markdown(
                    f"""
                    <div style="padding:1rem; border-radius:0.5rem; background-color:{color}20; margin-bottom:1rem;
                    border-left:0.5rem solid {color};">
                    <h4 style="margin:0; color:{color};">{rec['title']} <span style="float:right; font-size:0.8rem; 
                    background-color:{color}40; padding:0.25rem 0.5rem; border-radius:1rem;">{rec['priority']}</span></h4>
                    <p style="margin:0.5rem 0 0 0;">{rec['description']}</p>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
        else:
            st.info("Nenhuma recomendação específica para este projeto no momento.")
        
        # Fatores de risco
        risk_factors = analysis.get('risk_assessment', {}).get('factors', [])
        
        if risk_factors:
            st.markdown("### ⚠️ Fatores de Risco")
            
            for factor in risk_factors:
                color = '#F44336' if factor['severity'] == 'Alto' else '#FFC107' if factor['severity'] == 'Médio' else '#4CAF50'
                
                st.markdown(
                    f"""
                    <div style="padding:0.75rem; border-radius:0.5rem; background-color:{color}10; margin-bottom:0.5rem;
                    border-left:0.25rem solid {color};">
                    <h4 style="margin:0; font-size:1rem; color:{color};">{factor['factor']} 
                    <span style="float:right; font-size:0.7rem; background-color:{color}30; 
                    padding:0.15rem 0.3rem; border-radius:1rem;">{factor['severity']}</span></h4>
                    <p style="margin:0.3rem 0 0 0; font-size:0.9rem;">{factor['description']}</p>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
        
    except Exception as e:
        st.error(f"Erro ao renderizar análise de IA: {str(e)}")
        st.warning("A análise inteligente do projeto não está disponível no momento.")