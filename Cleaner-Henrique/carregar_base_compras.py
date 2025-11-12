import pandas as pd
import sys
import re
import os

# --- FIX: Adiciona o diretório do script ao path ---
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)
# --- FIM DO FIX ---

from firebase_utils import get_db

# --- Configurações ---
NOME_DA_PLANILHA_EXCEL = 'Dados' 
COLUNA_DE_VALOR = 'Vlr.Total'
COLECAO_FIRESTORE = 'base_compras' # Nome da nossa coleção no Firebase

# --- Função de Limpeza de Moeda BRL (Sem alterações) ---
def to_number_brl(x):
    try:
        if pd.isna(x): return 0.0
        s = str(x).strip().replace(' ', '') 
        if s == "": return 0.0
        
        if s.count(',') == 1 and s.count('.') > 0:
            s = s.replace('.', '').replace(',', '.')
        elif s.count(',') == 1:
            s = s.replace(',', '.')
        elif s.count('.') == 1 and s.count(',') == 0:
            pass 
        elif s.count('.') > 0 and s.count(',') == 0:
            s = s.replace('.', '')
        
        return float(s)
    except Exception:
        return 0.0

# --- FUNÇÃO 1: Ler o Excel para um DataFrame (Lógica antiga) ---
def ler_excel_para_df(arquivo_excel_bytesio):
    """
    Lê um arquivo Excel (recebido do Streamlit) e retorna um DataFrame limpo.
    """
    print(f"Iniciando leitura do arquivo Excel da memória...")
    try:
        df_compras = pd.read_excel(arquivo_excel_bytesio, 
                                 sheet_name=NOME_DA_PLANILHA_EXCEL, 
                                 engine='openpyxl',
                                 dtype=str) 
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        raise e 

    print("Limpando nomes das colunas...")
    def limpar_nome_coluna(nome):
        return str(nome).strip().replace('/', '_')
    
    df_compras.columns = [limpar_nome_coluna(col) for col in df_compras.columns]
    
    col_vlr_limpo = limpar_nome_coluna(COLUNA_DE_VALOR)
    
    if col_vlr_limpo in df_compras.columns:
        print(f"Convertendo a coluna de valor BRL '{col_vlr_limpo}' para numérico...")
        df_compras[col_vlr_limpo] = df_compras[col_vlr_limpo].apply(to_number_brl)
    else:
        print(f"Aviso: A coluna de valor '{col_vlr_limpo}' não foi encontrada.")

    # Importante: Converte colunas que podem ser interpretadas como data para str
    # para evitar que o Firebase as salve no formato errado
    for col in df_compras.columns:
        if 'data' in col.lower():
            df_compras[col] = df_compras[col].astype(str)

    print(f"Leitura concluída. {len(df_compras)} linhas encontradas.")
    return df_compras

# --- FUNÇÃO 2: Fazer o Upload do DataFrame para o Firebase ---
def carregar_base_firebase(df, modo_execucao='replace'):
    """
    Carrega o DataFrame para a coleção 'base_compras' no Firestore.
    """
    db = get_db()
    if db is None:
        raise Exception("Não foi possível conectar ao Firestore.")
    
    # --- Lógica de "Replace" ---
    if modo_execucao == 'replace':
        print("--- MODO REPLACE ---")
        print(f"Apagando todos os documentos da coleção '{COLECAO_FIRESTORE}'...")
        
        # Deletar em lotes (obrigatório pelo Firestore)
        docs = db.collection(COLECAO_FIRESTORE).limit(500).stream()
        deleted = 0
        
        while True:
            batch = db.batch()
            doc_count = 0
            for doc in docs:
                batch.delete(doc.reference)
                doc_count += 1
                deleted += 1
            
            if doc_count == 0:
                break # Sai do loop se não houver mais docs
            
            batch.commit()
            print(f"Lote de {doc_count} documentos apagado...")
            # Pega o próximo lote
            docs = db.collection(COLECAO_FIRESTORE).limit(500).stream()

        print(f"Total de {deleted} documentos antigos apagados.")
    
    print(f"Iniciando upload de {len(df)} novos registros para o Firebase...")
    
    # Converte o DF para uma lista de dicionários
    # fillna('') para evitar problemas com valores NaN, que o Firestore não aceita
    registros = df.fillna('').to_dict('records')
    
    # --- Lógica de Upload em Lote (Batch) ---
    # O Firestore tem um limite de 500 operações por lote
    batch = db.batch()
    total_carregado = 0
    
    for i, record in enumerate(registros):
        # Cria uma nova referência de documento (com ID automático)
        doc_ref = db.collection(COLECAO_FIRESTORE).document()
        batch.set(doc_ref, record)
        
        total_carregado += 1
        
        # Faz o commit do lote a cada 499 registros (para segurança)
        if (i + 1) % 499 == 0:
            print(f"Commitando lote... {total_carregado} / {len(registros)} registros carregados.")
            batch.commit()
            # Inicia um novo lote
            batch = db.batch()

    # Commit do lote final (o que sobrou)
    batch.commit()
    print(f"Upload concluído! Total de {total_carregado} registros salvos no Firebase.")
    return True



