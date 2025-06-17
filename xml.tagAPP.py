import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from datetime import datetime

# --- Configuração da Página ---
st.set_page_config(
    page_title="Conversor XML para Excel", 
    layout="wide",
    page_icon="📄"
)

# --- Funções ---
def extrair_xmls_de_zip(zip_path, destino):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(destino)
    return [os.path.join(raiz, nome)
            for raiz, _, arquivos in os.walk(destino)
            for nome in arquivos if nome.lower().endswith('.xml')]

def formatar_valor(val):
    if val is None:
        return '0,00'
    return val.replace('.', ',')

# --- Processamento NFe ---
def processar_nfe_por_cabecalho(caminho_xml, ns):
    # (Mantenha a função processar_nfe original que já temos)
    # Retorna dados consolidados por CFOP
    pass

def processar_nfe_por_item(caminho_xml, ns):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        infNFe = root.find('.//ns:infNFe', ns)
        if infNFe is None:
            return []

        ide = infNFe.find('ns:ide', ns)
        emit = infNFe.find('ns:emit', ns)
        dest = infNFe.find('ns:dest', ns)
        itens = infNFe.findall('ns:det', ns)
        total = infNFe.find('ns:total/ns:ICMSTot', ns)
        infAdic = infNFe.find('ns:infAdic', ns)

        dados_itens = []
        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)
            
            # Dados básicos do item
            dados = {
                'Chave NFe': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                'Número NFe': ide.findtext('ns:nNF', default='', namespaces=ns),
                'Data Emissão': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
                'Emitente': emit.findtext('ns:xNome', default='', namespaces=ns),
                'CNPJ Emitente': emit.findtext('ns:CNPJ', default='', namespaces=ns),
                'Destinatário': dest.findtext('ns:xNome', default='', namespaces=ns) if dest is not None else '',
                'CNPJ Destinatário': dest.findtext('ns:CNPJ', default='', namespaces=ns) if dest is not None else '',
                'CFOP': prod.findtext('ns:CFOP', default='', namespaces=ns),
                'Código Produto': prod.findtext('ns:cProd', default='', namespaces=ns),
                'Descrição': prod.findtext('ns:xProd', default='', namespaces=ns),
                'NCM': prod.findtext('ns:NCM', default='', namespaces=ns),
                'Quantidade': prod.findtext('ns:qCom', default='', namespaces=ns),
                'Unidade': prod.findtext('ns:uCom', default='', namespaces=ns),
                'Valor Unitário': formatar_valor(prod.findtext('ns:vUnCom', default='0', namespaces=ns)),
                'Valor Total': formatar_valor(prod.findtext('ns:vProd', default='0', namespaces=ns)),
                'Observações': infAdic.findtext('ns:infCpl', default='', namespaces=ns) if infAdic is not None else ''
            }

            # Adiciona informações de impostos
            if imposto is not None:
                icms = imposto.find('.//ns:ICMS', ns)
                if icms is not None:
                    for child in icms:
                        if child.tag.endswith('ICMS00') or child.tag.endswith('ICMS20'):  # Adapte para outros tipos de ICMS
                            dados['Alíquota ICMS'] = child.findtext('ns:pICMS', default='', namespaces=ns)
                            dados['Valor ICMS'] = formatar_valor(child.findtext('ns:vICMS', default='0', namespaces=ns))
                            break

            dados_itens.append(dados)

        return dados_itens

    except Exception as e:
        st.error(f"Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

# --- Interface ---
def main():
    st.title("Conversor XML para Excel")
    st.caption("Desenvolvido por Patricia Gutierrez")
    
    tipo_doc = st.radio(
        "Selecione o tipo de documento:", 
        ["NFe", "CTe"], 
        horizontal=True
    )
    
    # Mostrar opção de layout apenas para NFe
    if tipo_doc == "NFe":
        layout = st.radio(
            "Layout de saída:",
            ["Cabeçalho", "Por Item"],
            horizontal=True
        )
    
    uploaded_file = st.file_uploader(
        "Selecione o arquivo ZIP com os XMLs", 
        type="zip"
    )

    if uploaded_file:
        with st.spinner("Processando arquivos..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_path = os.path.join(temp_dir, uploaded_file.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                xml_files = extrair_xmls_de_zip(zip_path, temp_dir)
                
                if not xml_files:
                    st.warning("Nenhum arquivo XML encontrado no ZIP.")
                else:
                    st.info(f"{len(xml_files)} arquivo(s) encontrado(s)")
                    
                    progress_bar = st.progress(0)
                    dados_totais = []
                    
                    for i, xml_file in enumerate(xml_files):
                        progress_bar.progress((i + 1) / len(xml_files))
                        
                        if tipo_doc == "NFe":
                            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                            if layout == "Cabeçalho":
                                dados_totais.extend(processar_nfe_por_cabecalho(xml_file, ns))
                            else:
                                dados_totais.extend(processar_nfe_por_item(xml_file, ns))
                        else:
                            dados_totais.extend(processar_cte(xml_file))
                    
                    if dados_totais:
                        df = pd.DataFrame(dados_totais)
                        st.dataframe(df)
                        
                        excel_data = criar_excel(df)
                        nome_arquivo = f"{tipo_doc}_{layout}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        
                        st.download_button(
                            "Baixar Excel",
                            data=excel_data,
                            file_name=nome_arquivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

def criar_excel(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            data = f.read()
        os.unlink(tmp.name)
    return data

if __name__ == "__main__":
    main()
