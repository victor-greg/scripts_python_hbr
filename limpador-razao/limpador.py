import streamlit as st
import pandas as pd
import re
import io
import os  
import xlsxwriter 
import xml.etree.ElementTree as ET 
from datetime import datetime

# --- OBTÉM O CAMINHO DO SCRIPT (PARA ACHAR OS ASSETS) ---
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()

# --- CAMINHOS DOS ASSETS ---
LOGO_PATH = os.path.join(SCRIPT_DIR, "assets", "logo.png") 

# --- (Suas funções de lógica permanecem 100% iguais) ---

def _read_xml_with_elementtree(uploaded_file):
    """
    Lê o arquivo XML SpreadsheetML usando ElementTree para
    converter a planilha '3-Lançamentos Contábeis' em um DataFrame.
    """
    try:
        # Resetar o ponteiro do arquivo, caso tenha sido lido antes
        uploaded_file.seek(0)
        tree = ET.parse(uploaded_file)
        root = tree.getroot()

        # Namespaces são cruciais para encontrar os elementos
        ns_map = {
            'd': 'urn:schemas-microsoft-com:office:spreadsheet',
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
        }
        
        ws_name = '3-Lançamentos Contábeis'
        worksheet = None
        for ws in root.findall('d:Worksheet', ns_map):
            if ws.attrib.get(f'{{{ns_map["ss"]}}}Name') == ws_name:
                worksheet = ws
                break
        
        if worksheet is None:
            st.error(f"Erro Crítico: Não foi possível encontrar a planilha '{ws_name}' no XML.")
            return None

        table = worksheet.find('d:Table', ns_map)
        if table is None:
            st.error("Erro Crítico: Estrutura XML inválida. Tag <Table> não encontrada.")
            return None
            
        rows = table.findall('d:Row', ns_map)
        
        if len(rows) < 2:
            st.error("Erro Crítico: Planilha não contém linhas de cabeçalho ou dados.")
            return None

        headers = []
        header_cells = rows[1].findall('d:Cell', ns_map)
        for cell in header_cells:
            data = cell.find('d:Data', ns_map)
            if data is not None and data.text is not None:
                headers.append(data.text)
            else:
                headers.append(f"Coluna_Vazia_{len(headers)}")

        data_list = []
        for row_elem in rows[2:]:
            cells = row_elem.findall('d:Cell', ns_map)
            if not cells: continue
                
            row_data = {}
            for i, cell in enumerate(cells):
                if i >= len(headers): break
                
                data_elem = cell.find('d:Data', ns_map)
                text = data_elem.text if data_elem is not None else None
                
                data_type = 'String'
                if data_elem is not None:
                    data_type = data_elem.attrib.get(f'{{{ns_map["ss"]}}}Type', 'String')
                
                val = text
                if text is not None:
                    if data_type == 'DateTime':
                        val = pd.to_datetime(text) 
                    elif data_type == 'Number':
                        val = pd.to_numeric(text, errors='coerce')
                
                row_data[headers[i]] = val
            
            if row_data:
                data_list.append(row_data)

        if not data_list:
            st.error("Nenhum dado encontrado nas linhas da planilha.")
            return None
            
        df = pd.DataFrame(data_list)
        return df

    except ET.ParseError as e:
        st.error(f"Erro ao processar o XML: {e}")
        return None
    except Exception as e:
        st.error(f"Um erro inesperado ocorreu during a leitura do XML com ElementTree:")
        st.exception(e)
        return None


def processar_arquivo_xml(uploaded_file):
    """
    Função principal para ler, processar e estilizar o arquivo XML/Excel.
    """
    
    df = _read_xml_with_elementtree(uploaded_file)
    
    if df is None:
        return None

    # --- 2. LIMPEZA: JUNÇÃO DE LINHAS DE HISTÓRICO ---
    processed_rows = []
    last_valid_row = None

    for _, row in df.iterrows():
        if pd.isna(row['LOTE/SUB/DOC/LINHA']) or row['LOTE/SUB/DOC/LINHA'] == '':
            if last_valid_row is not None:
                hist_atual = str(row['HISTORICO']).strip() if pd.notna(row['HISTORICO']) else ""
                last_valid_row['HISTORICO'] += f" {hist_atual}"
        else:
            if last_valid_row is not None:
                processed_rows.append(last_valid_row)
            last_valid_row = row.to_dict()
    
    if last_valid_row is not None:
        processed_rows.append(last_valid_row)

    df_processed = pd.DataFrame(processed_rows)
    
    if df_processed.empty:
        st.error("O processamento não gerou dados. Verifique o conteúdo da planilha.")
        return None

    # --- 3. CRIAÇÃO DA COLUNA "DOC" ---
    prefixes = ['NF.:', 'DOC.:', 'NF:', 'DOC:', 'TIT:', 'TIT.:', 'DUPL.:']
    regex_pattern = rf"(?:{'|'.join(prefixes)})\s*([0-9]{{6,9}})"
    
    df_processed['DOC'] = df_processed['HISTORICO'].astype(str).str.extract(
        regex_pattern, 
        flags=re.IGNORECASE
    ).fillna('')

    # --- 4. CRIAÇÃO DA COLUNA "CRED/DEB" ---
    df_processed['DEBITO'] = pd.to_numeric(df_processed['DEBITO'], errors='coerce').fillna(0)
    df_processed['CREDITO'] = pd.to_numeric(df_processed['CREDITO'], errors='coerce').fillna(0)

    def calcular_cred_deb(row):
        if row['DEBITO'] != 0:
            return row['DEBITO'] * -1
        elif row['CREDITO'] != 0:
            return row['CREDITO']
        else:
            return 0

    df_processed['CRED/DEB'] = df_processed.apply(calcular_cred_deb, axis=1)

    df_processed['DEBITO'] = df_processed['DEBITO'].round(2)
    df_processed['CREDITO'] = df_processed['CREDITO'].round(2)
    df_processed['CRED/DEB'] = df_processed['CRED/DEB'].round(2)

    # --- 5. REORDENAÇÃO DAS COLUNAS ---
    cols = list(df_processed.columns)
    
    if 'DOC' in cols: cols.remove('DOC')
    if 'CRED/DEB' in cols: cols.remove('CRED/DEB')
    
    if 'HISTORICO' in cols:
        cols.insert(cols.index('HISTORICO') + 1, 'DOC')
    else:
        cols.append('DOC')

    if 'CREDITO' in cols:
        cols.insert(cols.index('CREDITO') + 1, 'CRED/DEB')
    else:
        cols.append('CRED/DEB')

    df_final = df_processed[cols]

    return df_final

def criar_excel_estilizado(df):
    """
    Cria um arquivo Excel .xlsx em memória com toda a formatação solicitada.
    """
    output = io.BytesIO()
    
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    sheet_name = 'Lançamentos_Processados'
    
    df_table = df.copy()
    
    if 'DATA' in df_table.columns:
        df_table['DATA'] = pd.to_datetime(
            df_table['DATA'], errors='coerce'
        ).dt.date
    
    workbook = writer.book
    
    font_base = {'font_name': 'Courier New', 'font_size': 10}
    note_bg = '#FFFFE0' 
    acc_fmt_str = '#.##0,00;-#.##0,00;0,00' 
    
    text_format = workbook.add_format({**font_base, 'num_format': '@'})
    date_format = workbook.add_format({**font_base, 'num_format': 'dd/mm/yyyy'})
    acc_format = workbook.add_format({**font_base, 'num_format': acc_fmt_str})
    
    note_text_format = workbook.add_format({**font_base, 'num_format': '@', 'bg_color': note_bg})
    note_acc_format = workbook.add_format({**font_base, 'num_format': acc_fmt_str, 'bg_color': note_bg})
    
    column_settings = []
    for col_name in df_table.columns:
        if col_name == 'DATA':
            fmt = date_format
        elif col_name in ['DEBITO', 'CREDITO']:
            fmt = acc_format
        elif col_name == 'DOC':
            fmt = note_text_format
        elif col_name == 'CRED/DEB':
            fmt = note_acc_format
        else:
            fmt = text_format 
            
        column_settings.append({'header': col_name, 'format': fmt})

    worksheet = workbook.add_worksheet(sheet_name)
    (max_row, max_col) = df_table.shape
    data_list = df_table.where(pd.notna(df_table), None).values.tolist()

    worksheet.add_table(0, 0, max_row, max_col - 1, {
        'data': data_list,
        'columns': column_settings,
        'style': 'Table Style Medium 9' 
    })
    
    for i, col in enumerate(df_table.columns):
        header_len = len(str(col))
        data_len = df_table[col].astype(str).str.len().max()
        max_len = max(header_len, data_len) + 2
        worksheet.set_column(i, i, max_len)

    writer.close()
    output.seek(0)
    
    return output

# --- FIM DAS FUNÇÕES DE LÓGICA ---


# ==========================================================
# --- INTERFACE DO STREAMLIT (LÓGICA DOS BOTÕES) ---
# ==========================================================

st.set_page_config(layout="wide", page_title="Limpador de XML Contábil")

# --- 1. INICIALIZA O ESTADO (SESSION STATE) ---
# Isso é o que "lembra" em qual página estamos.
if 'app_mode' not in st.session_state:
    st.session_state.app_mode = "Limpador de XML" # Página padrão

# --- 2. BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.title("Limpador de Razão")
    
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_column_width=True)
    else:
        st.warning("Logo não encontrada em 'assets/logo.png'.")
        
    st.divider()

    # --- O menu de navegação com BOTÕES ---
    # use_container_width=True faz os botões ocuparem a largura da sidebar
    
    st.header("Navegação")
    if st.button("Limpador de XML", use_container_width=True):
        st.session_state.app_mode = "Limpador de XML"

    if st.button("Outra Funcionalidade (Futuro)", use_container_width=True):
        st.session_state.app_mode = "Outra Funcionalidade (Futuro)"

    if st.button("Sobre", use_container_width=True):
        st.session_state.app_mode = "Sobre"
    
    st.divider()
    st.write("Versão 1.0.10") # <-- A versão do app

# --- 3. CONTEÚDO PRINCIPAL (BASEADO NO ESTADO) ---
# O app verifica o valor em st.session_state.app_mode para decidir o que mostrar

if st.session_state.app_mode == "Limpador de XML":
    
    # --- Coloquei o código antigo da sua UI aqui ---
    st.title("Ferramenta de Limpeza de XML Contábil")
    st.markdown("**Faça o upload do seu arquivo XML (formato Excel) para processamento.**")
    st.divider()

    uploaded_file = st.file_uploader(
        "Selecione o arquivo XML (exportado pelo Protheus)", 
        type=["xml", "xls", "xlsx"]
    )

    if uploaded_file:
        with st.spinner("Processando o arquivo com ElementTree... ⚙️"):
            df_final = processar_arquivo_xml(uploaded_file)
        
        if df_final is not None:
            st.success("Arquivo processado com sucesso! 🎉")
            
            st.subheader("Prévia dos Dados Processados")
            st.dataframe(df_final.head(50))

            st.subheader("Download do Arquivo Limpo")
            st.info("O arquivo abaixo está no formato .xlsx e contém todas as formatações solicitadas.")
            
            with st.spinner("Gerando arquivo Excel estilizado... 🎨"):
                excel_data = criar_excel_estilizado(df_final)
            
            # Gera o nome do novo arquivo
            original_name = os.path.splitext(uploaded_file.name)[0]
            new_filename = f"LIMPADO_{original_name}.xlsx"
            
            st.download_button(
                label="Clique aqui para baixar o .xlsx Processado",
                data=excel_data,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- Exemplo de como adicionar uma nova "janela" ---
elif st.session_state.app_mode == "Outra Funcionalidade (Futuro)":
    st.title("Outra Funcionalidade 🚀")
    st.write("Esta página está em construção.")
    st.info("Quando você quiser criar uma nova ferramenta, basta adicioná-la aqui.")

elif st.session_state.app_mode == "Sobre":
    st.title("Sobre o App")
    st.write("Este aplicativo foi criado para limpar e formatar arquivos XML contábeis exportados do TOTVS.")
    st.write("Ele usa o ElementTree para ler o XML de forma robusta e o XlsxWriter para criar o arquivo Excel de saída formatado.")
