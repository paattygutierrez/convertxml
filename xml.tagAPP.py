import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import tempfile
import streamlit as st

# Configuração inicial
st.set_page_config(
    page_title="Conversor XML NFe para Excel",
    layout="wide",
    page_icon="📄",
    initial_sidebar_state="expanded"
)

# Função para criar Excel (deve vir antes de ser usada!)
def criar_excel(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            bytes_data = f.read()
        os.unlink(tmp.name)
    return bytes_data

def extrair_xmls_de_zip(zip_path, destino):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(destino)
    return [os.path.join(raiz, nome) 
            for raiz, _, arquivos in os.walk(destino) 
            for nome in arquivos 
            if nome.lower().endswith('.xml')]

def processar_xml(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
        
        infNFe = root.find('.//ns:infNFe', ns)
        if infNFe is None:
            return None
        
        chave = infNFe.attrib.get('Id', '')[3:]
        numero = root.findtext('.//ns:nNF', namespaces=ns)
        emitente = root.findtext('.//ns:xNome', namespaces=ns)
        
        return {'Chave': chave, 'Número': numero, 'Emitente': emitente}
    
    except Exception as e:
        st.error(f"Erro no arquivo {os.path.basename(caminho_xml)}: {str(e)}")
        return None

# Interface do usuário
st.title("📄 Conversor XML NFe para Excel")
st.markdown("**Desenvolvido por Patricia Gutierrez**")
st.write("Esta aplicação converte arquivos XML de NFe em planilhas Excel.")

uploaded_file = st.file_uploader("Selecione o arquivo ZIP com os XMLs", type="zip")

if uploaded_file:
    with st.spinner('Processando arquivos...'):
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, uploaded_file.name)
            with open(zip_path, "wb") as f:
                f.write(uploaded_file.read())

            xml_files = extrair_xmls_de_zip(zip_path, temp_dir)
            if not xml_files:
                st.warning("Nenhum XML encontrado no arquivo ZIP!")
            else:
                st.success(f"{len(xml_files)} arquivos XML encontrados")

                dados = []
                progress_bar = st.progress(0)
                for i, xml_file in enumerate(xml_files):
                    progress_bar.progress((i + 1) / len(xml_files))
                    result = processar_xml(xml_file)
                    if result:
                        dados.append(result)

                if dados:
                    df = pd.DataFrame(dados)
                    st.dataframe(df)

                    excel_buffer = criar_excel(df)
                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=excel_buffer,
                        file_name="NFes_exportadas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
