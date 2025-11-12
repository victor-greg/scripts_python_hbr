# streamlit_app.py
import streamlit as st
import pandas as pd
import tempfile
import os
import io
import sys
from PIL import Image 

# --- Adiciona o path e importa as funções ---
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

from carregar_base_compras import ler_excel_para_df, carregar_base_firebase
from rodar_conciliacao import rodar_conciliacao_streamlit
# --- NOVA IMPORTAÇÃO ---
from firebase_utils import get_db, query_base_compras

# --- Caminhos dos Assets ---
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
    st.write("Versão 2.1 (com Auditoria)")

# --- Título Principal ---
st.title("🚀 Conciliador de Títulos TOTVS")
st.markdown("Uma ferramenta eficiente para a gestão financeira (com banco de dados persistente).")
st.markdown("---")

# --- Inicialização do Session State ---
if 'download_data' not in st.session_state:
    st.session_state.download_data = None
if 'download_filename' not in st.session_state:
    st.session_state.download_filename = None
if 'df_audit_cache' not in st.session_state:
    st.session_state.df_audit_cache = pd.DataFrame() # Cache para os dados da auditoria

# --- ABAS DA APLICAÇÃO ---
tab_conciliador, tab_auditoria = st.tabs(["🚀 Conciliador", "🔍 Auditoria da Base"])

# --- ABA 1: CONCILIADOR (SEU CÓDIGO ANTIGO) ---
with tab_conciliador:
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
                        st.stop()
                
                with st.spinner(f"Carregando {len(df_novo)} registros para o Firebase (Modo: {modo})... Isso pode levar vários minutos!"):
                    try:
                        carregar_base_firebase(df_novo, modo_execucao=modo)
                        st.success(f"🎉 Base de dados salva no Firebase com sucesso!")
                        # Limpa o cache da auditoria, pois os dados mudaram
                        st.session_state.df_audit_cache = pd.DataFrame()
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
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
                            tmp_xml.write(uploader_a.getvalue())
                            xml_path = tmp_xml.name
                        
                        excel_bytes_io, formatado_ok = rodar_conciliacao_streamlit(xml_path)
                        
                        st.session_state.download_data = excel_bytes_io
                        st.session_state.download_filename = "Relatorio_Final_Desmembrado.xlsx"
                        
                        st.success("🎉 Conciliação Concluída! O download está pronto.")
                        if not formatado_ok:
                            st.warning("Aviso: Falha na formatação avançada (OpenPyXL).")
                            
                    except Exception as e:
                        st.error(f"❌ Erro inesperado durante a conciliação:")
                        st.exception(e) 
                    finally:
                        if 'xml_path' in locals() and os.path.exists(xml_path):
                            os.remove(xml_path)
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
                on_click=lambda: st.session_state.update(download_data=None)
            )

# --- ABA 2: AUDITORIA (NOVO CÓDIGO) ---
with tab_auditoria:
    st.header("🔍 Auditoria e Visualização da Base de Compras")
    st.markdown("Busque e filtre os dados que estão salvos no banco de dados Firebase.")
    
    # --- Filtros do Firestore (Busca Rápida) ---
    st.subheader("1. Filtros Rápidos (Busca no Banco de Dados)")
    st.info("Estes filtros buscam no Firebase. São 'iguais a' e 'sensíveis a maiúsculas'. Deixe em branco para ignorar.")
    
    filt_col1, filt_col2, filt_col3 = st.columns(3)
    with filt_col1:
        f_fornecedor = st.text_input("Código Fornecedor (Ex: 000123)", key="f_forn")
    with filt_col2:
        f_documento = st.text_input("Nº Documento (Ex: 12345)", key="f_doc")
    with filt_col3:
        f_filial = st.text_input("Filial (Ex: 101)", key="f_filial")
        
    if st.button("BUSCAR DADOS DO FIREBASE", use_container_width=True, type="primary"):
        with st.spinner("Buscando dados no Firebase..."):
            try:
                # Chama a nova função de query
                df_audit = query_base_compras(
                    fornecedor=f_fornecedor if f_fornecedor else None,
                    documento=f_documento if f_documento else None,
                    filial=f_filial if f_filial else None
                )
                
                if df_audit.empty:
                    st.warning("Nenhum dado encontrado no Firebase com esses filtros.")
                    st.session_state.df_audit_cache = pd.DataFrame() # Limpa o cache
                else:
                    st.success(f"{len(df_audit)} registros baixados! Agora você pode usar os filtros adicionais.")
                    st.session_state.df_audit_cache = df_audit # Salva no cache
            except Exception as e:
                st.error(f"Erro ao buscar dados: {e}")
                st.session_state.df_audit_cache = pd.DataFrame()

    # --- Exibição e Filtros Locais (Pandas) ---
    if not st.session_state.df_audit_cache.empty:
        df = st.session_state.df_audit_cache.copy()
        st.markdown("---")
        st.subheader("2. Filtros Adicionais (sobre os dados baixados)")
        
        df_filtrado = df
        
        # --- Filtros de Range (Valor e Data) ---
        filt_col_range1, filt_col_range2 = st.columns(2)
        
        with filt_col_range1:
            st.markdown("**Filtrar por Valor (Vlr_Total)**")
            # Tenta converter Vlr_Total para numérico, tratando erros
            if 'Vlr_Total' in df.columns:
                df_filtrado['Vlr_Total'] = pd.to_numeric(df_filtrado['Vlr_Total'], errors='coerce').fillna(0.0)
                val_min = float(df_filtrado['Vlr_Total'].min())
                val_max = float(df_filtrado['Vlr_Total'].max())
                
                # Evita erro se min e max forem iguais
                if val_min == val_max:
                    f_valor_range = (val_min, val_max) # Não mostra o slider
                else:
                    f_valor_range = st.slider(
                        "Selecione o range de valor:",
                        min_value=val_min,
                        max_value=val_max,
                        value=(val_min, val_max),
                        key="f_valor_slider"
                    )
                
                df_filtrado = df_filtrado[
                    (df_filtrado['Vlr_Total'] >= f_valor_range[0]) &
                    (df_filtrado['Vlr_Total'] <= f_valor_range[1])
                ]
            else:
                st.warning("Coluna 'Vlr_Total' não encontrada para filtro de valor.")
        
        with filt_col_range2:
            st.markdown("**Filtrar por Data**")
            # Encontra colunas que parecem ser de data
            colunas_data_opcoes = [col for col in df.columns if 'data' in col.lower() or 'emissao' in col.lower() or 'vencto' in col.lower()]
            
            if colunas_data_opcoes:
                f_data_col = st.selectbox("Selecione a coluna de data para filtrar:", colunas_data_opcoes, key="f_data_col_select")
                
                # Converte a coluna de data (que é string) para datetime
                df_filtrado[f_data_col] = pd.to_datetime(df_filtrado[f_data_col], errors='coerce')
                
                # Remove NaT (datas inválidas) para o filtro
                df_sem_nat = df_filtrado.dropna(subset=[f_data_col])
                
                if not df_sem_nat.empty:
                    date_min = df_sem_nat[f_data_col].min().date()
                    date_max = df_sem_nat[f_data_col].max().date()

                    f_date_range = st.date_input(
                        "Selecione o range de data:",
                        value=(date_min, date_max),
                        min_value=date_min,
                        max_value=date_max,
                        key="f_date_range_picker"
                    )
                    
                    if len(f_date_range) == 2:
                        df_filtrado = df_filtrado[
                            (df_filtrado[f_data_col].dt.date >= f_date_range[0]) &
                            (df_filtrado[f_data_col].dt.date <= f_date_range[1])
                        ]
                else:
                    st.warning(f"Coluna '{f_data_col}' não contém datas válidas para filtro.")
            else:
                st.warning("Nenhuma coluna de data encontrada para filtro.")

        # --- Exibir DataFrame ---
        st.markdown("---")
        st.subheader(f"Dados Filtrados ({len(df_filtrado)} registros)")
        
        # Mostra o DF e permite que o usuário filtre as colunas
        st.dataframe(df_filtrado, use_container_width=True)
        
        # --- Botão de Download do CSV ---
        csv = df_filtrado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar CSV dos dados filtrados",
            data=csv,
            file_name="auditoria_base_compras.csv",
            mime="text/csv",
            use_container_width=True
        )
    elif 'df_audit_cache' in st.session_state and st.session_state.df_audit_cache.empty:
        # Se a busca foi feita mas não retornou nada
        st.info("A busca não retornou resultados. Tente filtros mais amplos (ou deixe em branco para buscar tudo).")
    else:
        # Estado inicial
        st.info("Clique em 'BUSCAR DADOS DO FIREBASE' para começar.")
        
st.markdown("---")
st.markdown("Desenvolvido com ❤️ para otimizar suas operações.")
