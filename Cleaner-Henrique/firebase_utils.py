import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore

def get_db():
    """
    Conecta-se ao Firebase e retorna uma instância do cliente Firestore.
    Usa st.secrets para autenticação e st.cache_resource para manter a conexão.
    """
    try:
        # Tenta obter a conexão do cache do Streamlit
        return st.cache_resource(init_firebase_app)()
    except Exception as e:
        st.error(f"Erro ao conectar ao Firebase: {e}")
        st.error("Verifique se você configurou os [firebase_service_account] corretamente nos Segredos do Streamlit (st.secrets).")
        return None

@st.cache_resource
def init_firebase_app():
    """
    Função de inicialização que será cacheada.
    """
    try:
        # Verifica se as credenciais estão nos segredos
        if "firebase_service_account" not in st.secrets:
            raise ValueError("Chave 'firebase_service_account' não encontrada nos st.secrets.")
        
        # Carrega as credenciais a partir dos segredos (que são um dict)
        creds_dict = dict(st.secrets["firebase_service_account"])
        
        # Inicializa o app Firebase
        cred = credentials.Certificate(creds_dict)
        
        # Evita re-inicialização
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)
            
        print("Conexão com Firebase estabelecida.")
        return firestore.client()
        
    except ValueError as e:
        # Erro comum se o TOML estiver mal formatado
        st.error(f"Erro ao carregar credenciais do Firebase: {e}")
        st.error("Verifique a formatação do seu TOML nos Segredos do Streamlit.")
        return None
    except Exception as e:
        st.error(f"Erro inesperado ao inicializar o Firebase: {e}")
        return None
# (No final de firebase_utils.py)
import pandas as pd

COLECAO_FIRESTORE = 'base_compras'

def query_base_compras(fornecedor=None, documento=None, filial=None):
    """
    Busca a base de compras no Firestore, aplicando filtros de IGUALDADE.
    """
    db = get_db()
    if db is None:
        raise Exception("Não foi possível conectar ao Firestore.")
    
    query = db.collection(COLECAO_FIRESTORE)
    
    # IMPORTANTE: Estes nomes de colunas ('Forn_Cliente', 'Documento', 'Filial')
    # devem ser os nomes exatos das chaves no seu Firestore (ou seja, os nomes
    # das colunas limpas do seu Excel). Ajuste se necessário.
    
    if fornecedor:
        # Assumindo que o nome da coluna no Firestore é 'Forn_Cliente'
        query = query.where('Forn_Cliente', '==', fornecedor)
    if documento:
        # Assumindo que o nome da coluna no Firestore é 'Documento'
        query = query.where('Documento', '==', documento)
    if filial:
        # Assumindo que o nome da coluna no Firestore é 'Filial'
        query = query.where('Filial', '==', filial)
    
    print(f"Executando query no Firestore (Fornecedor={fornecedor}, Documento={documento}, Filial={filial})...")
    docs = query.stream()
    dados = [doc.to_dict() for doc in docs]
    print(f"Query retornou {len(dados)} documentos.")
    
    if not dados:
        return pd.DataFrame()
        
    return pd.DataFrame(dados)
