import pandas as pd
import sqlite3
import sys
import os
import re

# --- Configurações ---
NOME_DA_PLANILHA_EXCEL = 'Dados' # 👈 Altere para o nome da aba correta
COLUNAS_CHAVE_EXCEL = {
    'codigo_fornecedor': 'Forn_Cliente', 
    'numero_documento': 'Documento',     
    'coluna_rateio': 'Rateio'            
}

COLUNA_DE_VALOR = 'Vlr.Total' # Use o nome *original* (antes da limpeza)

ARQUIVO_DB = 'base_compras.db'
NOME_TABELA = 'rateios_compras'

# --- (NOVO) Função de Limpeza de Moeda BRL ---
def to_number_brl(x):
    """Converte valores monetários BRL (com . e ,) para float."""
    try:
        if pd.isna(x): return 0.0
        s = str(x).strip().replace(' ', '') # Remove espaços
        if s == "": return 0.0
        
        # Formato: '40.891,47' (milhar . e decimal ,)
        if s.count(',') == 1 and s.count('.') > 0:
            s = s.replace('.', '').replace(',', '.')
        # Formato: '72,40' (só decimal ,)
        elif s.count(',') == 1:
            s = s.replace(',', '.')
        # Formato: '40891.47' (padrão US, já limpo)
        elif s.count('.') == 1 and s.count(',') == 0:
            pass # Já está correto
        # Formato '40.891' (milhar . sem decimal)
        elif s.count('.') > 0 and s.count(',') == 0:
            s = s.replace('.', '')
        
        return float(s)
    except Exception:
        return 0.0
# --- Fim da Função ---

def carregar_base_sqlite(caminho_arquivo_excel, modo_execucao='replace'):
    """
    Lê o arquivo Excel (caminho fornecido) e o salva no banco SQLite.
    FORÇA TODAS AS COLUNAS a serem lidas como TEXTO para preservar a formatação.
    """
    
    if not os.path.exists(caminho_arquivo_excel):
        print(f"Erro: Arquivo '{caminho_arquivo_excel}' não encontrado.")
        return False

    print(f"Iniciando leitura do '{caminho_arquivo_excel}' (Planilha: '{NOME_DA_PLANILHA_EXCEL}')...")
    
    # Força a leitura de TODAS as colunas como 'str'
    print("Forçando a leitura de todas as colunas como TEXTO...")
    try:
        df_compras = pd.read_excel(caminho_arquivo_excel, 
                                 sheet_name=NOME_DA_PLANILHA_EXCEL, 
                                 engine='openpyxl',
                                 dtype=str) # <-- FORÇA TUDO PARA TEXTO
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return False

    # --- LIMPEZA DOS CABEÇALHOS ---
    print("Limpando nomes das colunas (removendo espaços e trocando '/') ...")
    
    def limpar_nome_coluna(nome):
        return str(nome).strip().replace('/', '_')
    
    df_compras.columns = [limpar_nome_coluna(col) for col in df_compras.columns]
    print("Nomes das colunas limpos.")

    # --- (ATUALIZADO) Converter a coluna de valor para NÚMERO ---
    col_vlr_limpo = limpar_nome_coluna(COLUNA_DE_VALOR)
    
    if col_vlr_limpo in df_compras.columns:
        print(f"Convertendo a coluna de valor BRL '{col_vlr_limpo}' para numérico...")
        # Usa a nova função de limpeza
        df_compras[col_vlr_limpo] = df_compras[col_vlr_limpo].apply(to_number_brl)
    else:
        print(f"Aviso: A coluna de valor '{col_vlr_limpo}' (original: '{COLUNA_DE_VALOR}') não foi encontrada.")

    print(f"Leitura concluída. {len(df_compras)} linhas encontradas.")
    print(f"Conectando ao banco de dados SQLite '{ARQUIVO_DB}'...")
    
    conn = sqlite3.connect(ARQUIVO_DB)
    
    print(f"Salvando dados na tabela '{NOME_TABELA}' (modo: {modo_execucao})...")
    
    df_compras.to_sql(NOME_TABELA, conn, if_exists=modo_execucao, index=False)
    
    if modo_execucao == 'replace':
        print("Criando índices para otimizar consultas...")
        col_forn_db = COLUNAS_CHAVE_EXCEL['codigo_fornecedor']
        col_doc_db = COLUNAS_CHAVE_EXCEL['numero_documento']
        
        cursor = conn.cursor()
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_forn_doc ON {NOME_TABELA} ({col_forn_db}, {col_doc_db})")
        print(f"Índice 'idx_forn_doc' criado com sucesso.")
    
    conn.commit()
    conn.close()
    
    print("\n--- SUCESSO! ---")
    print(f"O banco de dados '{ARQUIVO_DB}' foi criado/atualizado.")
    return True