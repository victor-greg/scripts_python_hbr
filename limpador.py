import streamlit as st
import pandas as pd
import re
import io
import os  # Importado para verificar o caminho da logo
import xml.etree.ElementTree as ET 
from datetime import datetime

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
        
        # 1. Encontra a planilha correta pelo nome
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

        # 2. Pega os cabeçalhos da segunda linha (index 1)
        headers = []
        header_cells = rows[1].findall('d:Cell', ns_map)
        for cell in header_cells:
            data = cell.find('d:Data', ns_map)
            if data is not None and data.text is not None:
                headers.append(data.text)
            else:
                headers.append(f"Coluna_Vazia_{len(headers)}") # Fallback

        # 3. Processa as linhas de dados (começando do index 2)
        data_list = []
        for row_elem in rows[2:]:
            cells = row_elem.findall('d:Cell', ns_map)
            if not cells:
                continue # Pula linhas em branco
                
            row_data = {}
            for i, cell in enumerate(cells):
                if i >= len(headers):
                    break # Mais células que cabeçalhos, ignora
                
                data_elem = cell.find('d:Data', ns_map)
                text = data_elem.text if data_elem is not None else None
                
                # Pega o tipo de dado do XML para conversão correta
                data_type = 'String' # Default
                if data_elem is not None:
                    # O tipo está no atributo 'ss:Type'
                    data_type = data_elem.attrib.get(f'{{{ns_map["ss"]}}}Type', 'String')
                
                val = text
                if text is not None: # Só processa se houver texto
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
        st.info("O arquivo pode estar corrompido ou não ser um XML válido (SpreadsheetML).")
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

    # --- CORREÇÃO AQUI: Arredondar valores para 2 casas decimais ---
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
    note_bg = '#FFFFE0' # Amarelo (Estilo "Nota")
    
    # --- CORREÇÃO AQUI: Removido o '[Red]' do formato negativo ---
    acc_fmt_str = '#.##0,00;-#.##0,00;0,00' # Formato Contábil BR (Negativo em preto)
    
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

# --- INTERFACE DO STREAMLIT ---

st.set_page_config(layout="wide", page_title="Limpador de XML Contábil")

col1, col2 = st.columns([1, 5])
with col1:
    # --- CORREÇÃO AQUI: Caminho da logo ---
    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        st.image(logo_path, width=100)
    else:
        st.image(
            "https://upload.wikimedia.org/wikipedia/commons/c/c2/Office_apps_logo_2019.svg",
            width=100
        )
        st.warning("Logo não encontrada em 'assets/logo.png'.")

with col2:
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
