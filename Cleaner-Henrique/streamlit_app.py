# streamlit_app.py
import streamlit as st
import pandas as pd
import tempfile
import os
import io
from PIL import Image # Para manipular a imagem da logo

# Importa as funções de lógica refatoradas
from carregar_base_compras import preparar_base_compras
from rodar_conciliacao import rodar_conciliacao_streamlit

# --- Configurações da Página (AGORA COM FAVICON!) ---
# Caminhos para os assets
LOGO_PATH = "assets/logo_empresa.png" # Verifique se o nome do arquivo está correto
FAVICON_PATH = "assets/favicon.png"   # Verifique se o nome do arquivo está correto

# Verifica se os arquivos existem antes de tentar usá-los
favicon_icon = FAVICON_PATH if os.path.exists(FAVICON_PATH) else None

st.set_page_config(
    page_title="Conciliador de Títulos TOTVS",
    layout="wide",
    initial_sidebar_state="expanded", # Deixa a barra lateral aberta por padrão
    page_icon=favicon_icon # Define o ícone da página/aba
)

# --- Estilos CSS Personalizados para uma UI mais moderna ---
# Adiciona um toque de gradiente no título e ajusta a sidebar
st.markdown("""
    <style>
    .reportview-container .main .block-container{
        padding-top: 2rem;
        padding-right: 1rem;
        padding-left: 1rem;
        padding-bottom: 2rem;
    }
    .css-1d391kg { # Target the main content area for a bit more padding
        padding-left: 1rem;
        padding-right: 1rem;
    }
    h1 {
        font-size: 3em;
        color: #2e6c80; /* Um azul mais escuro */
        background: -webkit-linear-gradient(left, #2e6c80, #5ac8e6); /* Gradiente */
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: bold;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }
    h2 {
        color: #3f7b9e; /* Um tom de azul para subtítulos */
    }
    h3 {
        color: #4a90b4; /* Outro tom de azul */
    }
    .stButton>button {
        background-color: #4CAF50; /* Green */
        color: white;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 8px;
        border: none;
        transition: background-color 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .stFileUploader label {
        font-weight: bold;
        color: #3f7b9e;
    }
    .stCheckbox span {
        font-size: 0.9em;
        color: #555;
    }
    /* Estilo para a barra lateral */
    [data-testid="stSidebar"] {
        background-color: #f0f2f6; /* Um cinza claro */
    }
    [data-testid="stSidebar"] .stImage {
        margin-top: -30px; /* Ajusta a posição da logo na sidebar */
        margin-bottom: 20px;
    }
    .stTextInput>div>div>input {
        border-radius: 8px;
        border: 1px solid #ccc;
        padding: 8px;
    }
    .stDownloadButton button {
        background-color: #007bff !important; /* Azul para o botão de download */
    }
    .stDownloadButton button:hover {
        background-color: #0056b3 !important;
    }
    </style>
    """, unsafe_allow_html=True)


# --- Barra Lateral (Sidebar) com Logo ---
with st.sidebar:
    if os.path.exists(LOGO_PATH):
        try:
            # Carrega a imagem e redimensiona para se ajustar melhor à sidebar
            logo = Image.open(LOGO_PATH)
            # Calcula a nova largura e altura mantendo a proporção
            width_percent = (100 / float(logo.size[0]))
            new_height = int((float(logo.size[1]) * float(width_percent)))
            # Ajusta a largura máxima para se adequar à sidebar
            st.image(logo, use_column_width='always', caption="Logo da Empresa")
        except Exception as e:
            st.warning(f"Não foi possível carregar a logo em '{LOGO_PATH}': {e}")
            st.header("Logo da Empresa") # Fallback se a imagem falhar
    else:
        st.header("Logo da Empresa") # Fallback se o arquivo não existir
    
    st.markdown("---")
    st.subheader("Sobre este App")
    st.info(
        "Este aplicativo foi desenvolvido para conciliar títulos TOTVS "
        "com base em sua planilha de compras. Siga os passos para gerar o relatório final."
    )
    st.markdown("---")
    st.write("Versão 1.0.0")

# --- Título Principal ---
st.title("🚀 Conciliador de Títulos TOTVS")
st.markdown("Uma ferramenta eficiente para a gestão financeira.")
st.markdown("---")

# Inicializa o 'session_state' para guardar nossa base de compras
if 'base_compras' not in st.session_state:
    st.session_state.base_compras = None
if 'download_data' not in st.session_state:
    st.session_state.download_data = None
if 'download_filename' not in st.session_state:
    st.session_state.download_filename = None

# --- Colunas da Interface ---
col1, col2 = st.columns(2)

# --- Coluna 1: Carregar Base (Arquivo B) ---
with col1:
    st.header("Passo 1: Carregar Base de Compras (XLSX)")
    st.markdown("Envie o arquivo Excel contendo sua base de compras. "
                "Esta será a referência para a conciliação.")
    
    uploader_b = st.file_uploader(
        "Selecione o arquivo XLSX da Base de Compras", 
        type="xlsx",
        key="uploader_base_compras" # Adicionado key para evitar avisos
    )
    
    modo_replace = st.checkbox(
        "Substituir base de dados em cache", 
        value=True,
        help="Se marcado, substitui a base na memória. Se desmarcado, adiciona os dados do novo arquivo à base existente.",
        key="checkbox_replace_mode"
    )
    
    if st.button("1. Carregar Base de Compras", use_container_width=True):
        if uploader_b:
            with st.spinner("Lendo e preparando a base de compras..."):
                try:
                    df_novo = preparar_base_compras(uploader_b)
                    
                    if modo_replace or st.session_state.base_compras is None:
                        st.session_state.base_compras = df_novo
                    else:
                        st.session_state.base_compras = pd.concat([st.session_state.base_compras, df_novo]).drop_duplicates(
                            subset=[
                                'Forn_Cliente', # Substitua pelas colunas que identificam um registro único na sua base de compras
                                'Documento',    # Adicione mais colunas se a combinação for necessária
                                'Vlr.Total'
                            ], keep='first'
                        )
                    
                    st.success(f"✅ Base de dados carregada com sucesso! "
                               f"{len(st.session_state.base_compras)} linhas totais em memória.")
                    st.dataframe(st.session_state.base_compras.head()) # Mostra um preview
                except Exception as e:
                    st.error(f"❌ Erro ao carregar a Base de Compras: {e}")
        else:
            st.warning("⚠️ Por favor, selecione um arquivo XLSX da Base de Compras antes de carregar.")

# --- Coluna 2: Rodar Conciliação (Arquivo A) ---
with col2:
    st.header("Passo 2: Rodar Conciliação (XML TOTVS)")
    st.markdown("Agora, envie o arquivo XML gerado pelo TOTVS para iniciar o processo de conciliação.")
    
    if st.session_state.base_compras is None:
        st.info("ℹ️ Aguardando a 'Base de Compras' ser carregada no Passo 1 para continuar...")
    else:
        st.success(f"👍 Base de Compras com {len(st.session_state.base_compras)} linhas pronta na memória para conciliação.")
        
        uploader_a = st.file_uploader(
            "Selecione o arquivo XML do TOTVS (Arquivo A)", 
            type="xml",
            key="uploader_xml_totvs"
        )
        
        if st.button("2. RODAR CONCILIAÇÃO", use_container_width=True, type="primary"):
            if uploader_a:
                with st.spinner("⚙️ Processando conciliação... (Isso pode levar um tempo)"):
                    try:
                        # Precisamos salvar o XML temporariamente porque a função
                        # read_spreadsheetml espera um caminho de arquivo.
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
                            tmp_xml.write(uploader_a.getvalue())
                            xml_path = tmp_xml.name
                        
                        # Pega o DF da base de compras da memória (session_state)
                        df_base_compras = st.session_state.base_compras.copy()
                        
                        # Chama a função de lógica refatorada
                        excel_bytes_io, colunas_formatadas = rodar_conciliacao_streamlit(xml_path, df_base_compras)
                        
                        st.session_state.download_data = excel_bytes_io
                        st.session_state.download_filename = "Relatorio_Final_Desmembrado.xlsx"
                        
                        st.success("🎉 Conciliação Concluída! O download está pronto.")
                        
                    except Exception as e:
                        st.error(f"❌ Erro inesperado durante a conciliação: {e}")
                        st.exception(e) # Mostra o traceback completo para depuração
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
                use_container_width=True
            )

st.markdown("---")
st.markdown("Desenvolvido com ❤️ para otimizar suas operações.")
