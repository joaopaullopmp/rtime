import streamlit as st
import pandas as pd
import hashlib
from db_utils import read_table_from_db, save_table_to_db, execute_query

class Auth:
    def __init__(self):
        self.table_name = 'utilizadores'
    
    def initialize_session(self):
        if 'logged_in' not in st.session_state:
            st.session_state.logged_in = False
        if 'user_info' not in st.session_state:
            st.session_state.user_info = None
    
    def login(self, email, password):
        if not email or not password:
            return False, "Por favor, preencha todos os campos."
        
        try:
            # Obter dados da tabela de usuários do banco
            users_df = read_table_from_db(self.table_name)
            
            # Verificar se o email existe
            if users_df.empty or email not in users_df['email'].values:
                return False, "Email não encontrado."
            
            # Obter dados do usuário
            user_data = users_df[users_df['email'] == email].iloc[0]
            
            # Verificar a senha
            password_hash = hashlib.sha256(str(password).encode()).hexdigest()
            if password_hash != user_data['password']:
                return False, "Senha incorreta."
            
            # Login bem-sucedido
            return True, user_data.to_dict()
        
        except Exception as e:
            print(f"Erro ao fazer login: {e}")
            return False, f"Erro de autenticação: {str(e)}"
    
    def register_user(self, data):
        try:
            # Obter dados atuais
            users_df = read_table_from_db(self.table_name)
            
            # Verificar se o email já existe
            if not users_df.empty and data['email'] in users_df['email'].values:
                return False, "Este email já está registrado."
            
            # Hash da senha
            data['password'] = hashlib.sha256(str(data['password']).encode()).hexdigest()
            
            # Criar novo ID
            if users_df.empty:
                new_id = 1
            else:
                new_id = users_df['user_id'].max() + 1
            data['user_id'] = new_id
            
            # Adicionar novo usuário
            new_user_df = pd.DataFrame([data])
            updated_df = pd.concat([users_df, new_user_df], ignore_index=True)
            
            # Salvar no banco
            save_table_to_db(updated_df, self.table_name)
            
            return True, "Usuário registrado com sucesso!"
        
        except Exception as e:
            print(f"Erro ao registrar usuário: {e}")
            return False, f"Erro ao registrar: {str(e)}"