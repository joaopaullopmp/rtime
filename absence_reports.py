import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
import plotly.figure_factory as ff
import plotly.express as px
from database_manager import DatabaseManager

def ausencias_report_page():
    """Página de relatório de mapa de ausências"""
    st.title("Mapa de Ausências")
    
    # Inicializar o gerenciador de banco de dados
    db_manager = DatabaseManager()
    
    # Carregar dados do banco de dados em vez de Excel
    absences_df = db_manager.query_to_df("SELECT * FROM absences")
    users_df = db_manager.query_to_df("SELECT * FROM utilizadores")
    
    # Verificar se há dados
    if absences_df.empty:
        st.warning("Não há dados de ausências disponíveis.")
        return
    
    # Converter datas para datetime
    absences_df['start_date'] = pd.to_datetime(absences_df['start_date'], errors='coerce')
    absences_df['end_date'] = pd.to_datetime(absences_df['end_date'], errors='coerce')
    
    # Selecionar período
    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox(
            "Mês",
            range(1, 13),
            format_func=lambda x: calendar.month_name[x],
            index=datetime.now().month - 1
        )
    with col2:
        ano = st.selectbox(
            "Ano",
            range(datetime.now().year - 1, datetime.now().year + 2),
            index=1
        )
    
    # Definir período
    primeiro_dia = datetime(ano, mes, 1)
    ultimo_dia = datetime(ano, mes, calendar.monthrange(ano, mes)[1], 23, 59, 59)
    
    # Filtrar ausências no período
    filtered_absences = absences_df[
        ((absences_df['start_date'] >= primeiro_dia) & (absences_df['start_date'] <= ultimo_dia)) |
        ((absences_df['end_date'] >= primeiro_dia) & (absences_df['end_date'] <= ultimo_dia)) |
        ((absences_df['start_date'] <= primeiro_dia) & (absences_df['end_date'] >= ultimo_dia))
    ]
    
    # Associar nomes de usuários
    user_names = {}
    for _, user in users_df.iterrows():
        user_id = user['user_id']
        user_names[user_id] = f"{user['First_Name']} {user['Last_Name']}"
    
    # Preparar dados para o gráfico de Gantt
    gantt_data = []
    
    for _, absence in filtered_absences.iterrows():
        try:
            user_id = absence['user_id']
            user_name = user_names.get(user_id, f"Usuário {user_id}")
            
            # Ajustar datas para ficarem dentro do período selecionado
            start_date = max(absence['start_date'], primeiro_dia)
            end_date = min(absence['end_date'], ultimo_dia)
            
            # Verificar tipo de ausência
            absence_type = str(absence.get('absence_type', 'Outro'))
            if pd.isna(absence_type):
                absence_type = 'Outro'
            
            # Cor baseada no tipo de ausência
            if 'férias' in absence_type.lower():
                color = '#4CAF50'  # Verde para férias
                legend_group = "Férias"
            elif 'feriado' in absence_type.lower():
                color = '#FFC107'  # Amarelo para feriados
                legend_group = "Feriado"
            elif 'licença' in absence_type.lower():
                color = '#F44336'  # Vermelho para licenças
                legend_group = "Licença"
            else:
                color = '#9C27B0'  # Roxo para outros tipos
                legend_group = "Outro"
            
            # Calcular duração em dias
            duration_days = (end_date - start_date).days + 1
            
            # Adicionar registro para o gráfico
            gantt_data.append({
                'Task': user_name,
                'Start': start_date,
                'Finish': end_date + timedelta(days=1),  # Adicionar um dia para incluir o último dia
                'Resource': absence_type,
                'Tipo': legend_group,
                'Status': absence.get('status', ''),
                'Duração (dias)': duration_days,
                'Descrição': absence.get('description', '')
            })
        except Exception as e:
            st.error(f"Erro ao processar ausência: {e}")
    
    # Criar DataFrame com os dados para o gráfico
    if gantt_data:
        df_gantt = pd.DataFrame(gantt_data)
        
        # Criar tabela de resumo
        st.subheader("Tabela de Ausências")
        
        # Preparar tabela de resumo
        table_df = df_gantt.copy()
        table_df['Início'] = table_df['Start'].dt.strftime('%d/%m/%Y')
        table_df['Fim'] = (table_df['Finish'] - timedelta(days=1)).dt.strftime('%d/%m/%Y')
        
        # Selecionar e renomear colunas para a tabela
        table_df = table_df[['Task', 'Resource', 'Início', 'Fim', 'Duração (dias)', 'Status', 'Descrição']]
        table_df = table_df.rename(columns={'Task': 'Usuário', 'Resource': 'Tipo de Ausência'})
        
        # Mostrar tabela com ordenação e filtros
        st.dataframe(table_df, use_container_width=True)
        
        # Criar gráfico de Gantt usando figure_factory
        fig = ff.create_gantt(
            df_gantt,
            colors={
                'Férias': '#4CAF50',
                'Feriado': '#FFC107',
                'Licença': '#F44336',
                'Outro': '#9C27B0'
            },
            index_col='Resource',
            group_tasks=True,
            show_colorbar=True,
            title=f'Mapa de Ausências - {calendar.month_name[mes]} {ano}'
        )
        
        # Personalizar layout
        fig.update_layout(
            autosize=True,
            height=600,
            xaxis_title='Data',
            yaxis_title='Colaborador',
            yaxis={'categoryorder': 'category ascending'},
            legend_title='Tipo de Ausência'
        )
        
        # Mostrar gráfico
        st.plotly_chart(fig, use_container_width=True)
        
        # Estatísticas
        st.subheader("Estatísticas de Ausência")
        
        col1, col2, col3 = st.columns(3)
        
        # Total de ausências no mês
        with col1:
            total_ausencias = len(filtered_absences)
            st.metric("Total de Ausências", total_ausencias)
        
        # Tipos de ausência
        with col2:
            tipos_ausencia = df_gantt['Tipo'].value_counts()
            st.write("Tipos de Ausência:")
            for tipo, count in tipos_ausencia.items():
                st.write(f"- {tipo}: {count}")
        
        # Colaboradores ausentes
        with col3:
            colaboradores_ausentes = df_gantt['Task'].nunique()
            total_colaboradores = len(users_df)
            st.metric(
                "Colaboradores com Ausência",
                colaboradores_ausentes,
                f"{colaboradores_ausentes/total_colaboradores*100:.1f}% do total"
            )
        
        # Botão para exportar dados
        if st.button("Exportar para Excel"):
            # Preparar DataFrame para exportação
            export_df = table_df.copy()
            
            # Salvar para Excel
            output_path = 'mapa_ausencias.xlsx'
            export_df.to_excel(output_path, index=False)
            
            # Disponibilizar para download
            with open(output_path, 'rb') as f:
                bytes_data = f.read()
                st.download_button(
                    label="Baixar Relatório Excel",
                    data=bytes_data,
                    file_name="mapa_ausencias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            st.success("Relatório exportado com sucesso!")
    else:
        st.info(f"Não há ausências registradas para {calendar.month_name[mes]} de {ano}.")