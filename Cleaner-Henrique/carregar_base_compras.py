import pandas as pd
import sys
import os
import re

# --- Configurações ---
NOME_DA_PLANILHA_EXCEL = 'Dados' # 👈 Altere para o nome da aba correta
COLUNA_DE_VALOR = 'Vlr.Total' # Use o nome *original* (antes da limpeza)

# --- Função de Limpeza de Moeda BRL (Sem alterações) ---
def to_number_brl(x):
    """Converte valores monetários BRL (com . e ,) para float."""
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

# --- FUNÇÃO ATUALIZADA ---
def preparar_base_compras(arquivo_excel_bytesio):
    """
    Lê um arquivo Excel (recebido do Streamlit) e retorna um DataFrame limpo.
    FORÇA TODAS AS COLUNAS a serem lidas como TEXTO primeiro.
    """
    
    print(f"Iniciando leitura do arquivo Excel da memória (Planilha: '{NOME_DA_PLANILHA_EXCEL}')...")
    
    # Força a leitura de TODAS as colunas como 'str'
    try:
        # Lê diretamente do objeto de arquivo do Streamlit
        df_compras = pd.read_excel(arquivo_excel_bytesio, 
                                 sheet_name=NOME_DA_PLANILHA_EXCEL, 
                                 engine='openpyxl',
                                 dtype=str) # <-- FORÇA TUDO PARA TEXTO
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        raise e # Repassa o erro para o Streamlit

    # --- LIMPEZA DOS CABEÇALHOS ---
    print("Limpando nomes das colunas...")
    
    def limpar_nome_coluna(nome):
        return str(nome).strip().replace('/', '_')
    
    df_compras.columns = [limpar_nome_coluna(col) for col in df_compras.columns]
    print("Nomes das colunas limpos.")

    # --- Converter a coluna de valor para NÚMERO ---
    col_vlr_limpo = limpar_nome_coluna(COLUNA_DE_VALOR)
    
    if col_vlr_limpo in df_compras.columns:
        print(f"Convertendo a coluna de valor BRL '{col_vlr_limpo}' para numérico...")
        df_compras[col_vlr_limpo] = df_compras[col_vlr_limpo].apply(to_number_brl)
    else:
        print(f"Aviso: A coluna de valor '{col_vlr_limpo}' (original: '{COLUNA_DE_VALOR}') não foi encontrada.")

    print(f"Leitura concluída. {len(df_compras)} linhas encontradas.")
    
    # --- NÃO HÁ MAIS SQLITE ---
    # O banco de dados agora é o DataFrame que estamos retornando
    

    return df_compras
