import streamlit as st
import pandas as pd
import tempfile
import os
import io
from PIL import Image 

# --- Importações de Lógica (Modificadas) ---
# Usamos o . para importação relativa (assumindo que __init__.py existe)
from .carregar_base_compras import ler_excel_para_df, carregar_base_firebase
from .rodar_conciliacao import rodar_conciliacao_streamlit
from .firebase_utils import get_db # Importamos só para verificar a conexão

# --- DEFINIÇÃO DE CAMINHO ABSOLUTO (A PROVA DE FALHAS) ---
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()

LOGO_PATH = os.path.join(SCRIPT_DIR, "assets", "logo.png") 
FAVICON_PATH = os.path.join(SCRIPT_DIR, "assets", "icone.ico")   

# --- Configurações da Página ---
favicon_icon = FAVICON_PATH if os.path.exists(FAVICON_PATH) else None
st.set_page_config(
    page_title="Conciliador de Títulos TOTVS",
    layout="wide",
    initial_sidebar_state="expanded", 
    page_icon=favicon_icon 
)

# --- Estilos CSS (Cole seu CSS personalizado aqui) ---
st.markdown("""
    <style>
    /* ... COLOQUE SEU CSS AQUI ... */
    
    /* Estilos da Sidebar (Exemplo) */
    [data-testid="stSidebar"] {
        color: #E0E0E0; /* CINZA-CLARO */
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #FFFFFF; /* BRANCO */
    }
    [data-testid="stSidebar"] .stInfo {
        color: #FFFFFF; /* BRANCO */
    }
    </style>
    """, unsafe_allow_html=True)

# --- Barra Lateral (Sidebar) com Logo ---
with st.sidebar:
    if os.path.exists(LOGO_PATH):
        try:
            logo = Image.open(LOGO_PATH)
            st.image(logo, use_column_width='always') 
        except Exception as e:
            st.warning(f"Erro ao carregar a logo: {e}")
    else:
        st.error(f"Logo não encontrada em: {LOGO_PATH}")
        st.header("Logo da Empresa")
    
    st.markdown("---")
    st.subheader("Sobre este App")
    st.info(
        "Este aplicativo foi desenvolvido para conciliar títulos TOTVS "
        "com base em sua planilha de compras salva no banco de dados em nuvem."
    )
    st.markdown("---")
    # Verifica a conexão com o Firebase e mostra um status
    if 'db_status' not in st.session_state:
        if get_db() is not None:
            st.session_state.db_status = "Conectado"
        else:
            st.session_state.db_status = "Falha na Conexão"
            
    if st.session_state.db_status == "Conectado":
        st.success("☁️ Conectado ao Firebase")
    else:
        st.error("❌ Falha na conexão com Firebase")
        
    st.markdown("---")
    st.subheader("Sobre este App")
    st.info(
        "Este aplicativo foi desenvolvido para conciliar títulos TOTVS "
        "com base em sua planilha de compras. Siga os passos para gerar o relatório final.")
    st.markdown("---")
    st.write("Versão 2.0 (Firebase)")

# --- Título Principal ---
st.title("🚀 Conciliador de Títulos TOTVS")
st.markdown("Uma ferramenta eficiente para a gestão financeira (com banco de dados persistente).")
st.markdown("---")

# --- Não usamos mais o session_state para os dados ---
# Apenas para o botão de download
if 'download_data' not in st.session_state:
    st.session_state.download_data = None
if 'download_filename' not in st.session_state:
    st.session_state.download_filename = None

# --- Colunas da Interface ---
col1, col2 = st.columns(2)

# --- Coluna 1: Carregar Base (Arquivo B) ---
with col1:
    st.header("Passo 1: Carregar Base de Compras (XLSX)")
    st.markdown("Envie o arquivo Excel para o banco de dados em nuvem (Firebase).")
    
    uploader_b = st.file_uploader(
        "Selecione o arquivo XLSX da Base de Compras", 
        type="xlsx",
        key="uploader_base_compras"
    )
    
    modo_replace = st.checkbox(
        "Substituir base de dados existente (Modo Replace)", 
        value=True,
        help="Se marcado, APAGA toda a base antiga no Firebase antes de carregar a nova. Se desmarcado, apenas ADICIONA os novos dados.",
        key="checkbox_replace_mode"
    )
    
    if st.button("1. CARREGAR PARA NUVEM", use_container_width=True):
        if uploader_b:
            modo = 'replace' if modo_replace else 'append'
            
            with st.spinner(f"Lendo Excel..."):
                try:
                    df_novo = ler_excel_para_df(uploader_b)
                    st.success(f"✅ Excel lido! {len(df_novo)} linhas prontas para upload.")
                except Exception as e:
                    st.error(f"❌ Erro ao ler o Excel: {e}")
                    st.stop() # Para a execução
            
            with st.spinner(f"Carregando {len(df_novo)} registros para o Firebase (Modo: {modo})... Isso pode levar vários minutos!"):
                try:
                    carregar_base_firebase(df_novo, modo_execucao=modo)
                    st.success(f"🎉 Base de dados salva no Firebase com sucesso!")
                    # Limpa o preview antigo, se houver
                    if 'df_preview' in st.session_state:
                        del st.session_state.df_preview
                except Exception as e:
                    st.error(f"❌ Erro ao carregar para o Firebase: {e}")
        else:
            st.warning("⚠️ Por favor, selecione um arquivo XLSX da Base de Compras antes de carregar.")

# --- Coluna 2: Rodar Conciliação (Arquivo A) ---
with col2:
    st.header("Passo 2: Rodar Conciliação (XML TOTVS)")
    st.markdown("Faça o upload do XML. O app irá baixar a base da nuvem e processar.")
    
    uploader_a = st.file_uploader(
        "Selecione o arquivo XML do TOTVS (Arquivo A)", 
        type="xml",
        key="uploader_xml_totvs"
    )
    
    if st.button("2. RODAR CONCILIAÇÃO", use_container_width=True, type="primary"):
        if uploader_a:
            with st.spinner("⚙️ Baixando base do Firebase e processando... (Isso pode levar um tempo)"):
                try:
                    # Salva o XML temporariamente
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
                        tmp_xml.write(uploader_a.getvalue())
                        xml_path = tmp_xml.name
                    
                    # --- CHAMADA MODIFICADA ---
                    # Não passamos mais o df_base_compras, a função busca sozinha
                    excel_bytes_io, formatado_ok = rodar_conciliacao_streamlit(xml_path)
                    
                    st.session_state.download_data = excel_bytes_io
                    st.session_state.download_filename = "Relatorio_Final_Desmembrado.xlsx"
                    
                    st.success("🎉 Conciliação Concluída! O download está pronto.")
                    if not formatado_ok:
                        st.warning("Aviso: Falha na formatação avançada (OpenPyXL).")
                        
                except Exception as e:
                    st.error(f"❌ Erro inesperado durante a conciliação:")
                    st.exception(e) # Mostra o traceback completo
                finally:
                    if 'xml_path' in locals() and os.path.exists(xml_path):
                        os.remove(xml_path) # Limpa o arquivo temporário
        else:
            st.warning("⚠️ Por favor, selecione o arquivo XML do TOTVS antes de rodar a conciliação.")

    # --- Botão de Download ---
    if st.session_state.download_data:
        st.markdown("---")
        st.success("Seu relatório está pronto para baixar!")
        st.download_button(
            label="📥 Baixar Relatório Final (.xlsx)",
            data=st.session_state.download_data,
            file_name=st.session_state.download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click=lambda: st.session_state.update(download_data=None) # Limpa após clicar
        )

st.markdown("---")
st.markdown("Desenvolvido com ❤️ para otimizar suas operações.")

