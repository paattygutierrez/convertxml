import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from datetime import datetime

# --- XML Processing Functions ---
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

def obter_status_nfe(root, ns):
    protNFe = root.find('.//ns:protNFe', ns)
    if protNFe is not None:
        cStat = protNFe.findtext('.//ns:cStat', namespaces=ns)
        xMotivo = protNFe.findtext('.//ns:xMotivo', namespaces=ns)
        return f"{cStat} - {xMotivo}" if cStat and xMotivo else ''
    return ''

def processar_nfe(caminho_xml, ns):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        infNFe = root.find('.//ns:infNFe', ns)
        if infNFe is None:
            return []

        ide = infNFe.find('ns:ide', ns)
        emit = infNFe.find('ns:emit', ns)
        itens = infNFe.findall('ns:det', ns)
        total = infNFe.find('ns:total/ns:ICMSTot', ns)

        dados_por_cfop = {}
        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)
            icms = imposto.find('.//ns:ICMS', ns) if imposto is not None else None

            cfop = prod.findtext('ns:CFOP', default='', namespaces=ns)
            vProd = prod.findtext('ns:vProd', default='0', namespaces=ns)
            vICMS = imposto.findtext('.//ns:vICMS', default='0', namespaces=ns) if imposto is not None else '0'
            pICMS = ''
            if icms is not None:
                for child in icms:
                    aliq = child.findtext('ns:pICMS', namespaces=ns)
                    if aliq:
                        pICMS = aliq
                        break

            if cfop not in dados_por_cfop:
                dados_por_cfop[cfop] = {
                    'Numero NFe': ide.findtext('ns:nNF', default='', namespaces=ns),
                    'Data Emissão': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
                    'CNPJ Emitente': emit.findtext('ns:CNPJ', default='', namespaces=ns),
                    'Emitente': emit.findtext('ns:xNome', default='', namespaces=ns),
                    'UF Emitente': emit.findtext('ns:enderEmit/ns:UF', default='', namespaces=ns),
                    'CFOP': cfop,
                    'Vlr Nota': 0.0,
                    'Vlr ICMS': 0.0,
                    'Aliquota ICMS': pICMS,
                    'Vlr IPI': 0.0,
                    'Vlr PIS': 0.0,
                    'Vlr COFINS': 0.0,
                    'VICMS ST': 0.0,
                    'Vlr Frete': 0.0,
                    'Vlr Seguro': 0.0,
                    'Chave de acesso': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                }

            dados_por_cfop[cfop]['Vlr Nota'] += float(vProd.replace(',', '.'))
            dados_por_cfop[cfop]['Vlr ICMS'] += float(vICMS.replace(',', '.'))

        for cfop_dado in dados_por_cfop.values():
            cfop_dado['Vlr IPI'] = formatar_valor(total.findtext('ns:vIPI', default='0', namespaces=ns))
            cfop_dado['Vlr PIS'] = formatar_valor(total.findtext('ns:vPIS', default='0', namespaces=ns))
            cfop_dado['Vlr COFINS'] = formatar_valor(total.findtext('ns:vCOFINS', default='0', namespaces=ns))
            cfop_dado['VICMS ST'] = formatar_valor(total.findtext('ns:vST', default='0', namespaces=ns))
            cfop_dado['Vlr Frete'] = formatar_valor(total.findtext('ns:vFrete', default='0', namespaces=ns))
            cfop_dado['Vlr Seguro'] = formatar_valor(total.findtext('ns:vSeg', default='0', namespaces=ns))
            cfop_dado['Vlr Nota'] = formatar_valor(str(cfop_dado['Vlr Nota']))
            cfop_dado['Vlr ICMS'] = formatar_valor(str(cfop_dado['Vlr ICMS']))

        return list(dados_por_cfop.values())

    except Exception as e:
        st.error(f"Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

def processar_cte(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}

        infCte = root.find('.//ns:infCte', ns)
        if infCte is None:
            return []

        dados = {
            'Número CTe': infCte.findtext('ns:ide/ns:nCT', namespaces=ns) or '',
            'Data Emissão': (infCte.findtext('ns:ide/ns:dhEmi', namespaces=ns) or '')[:10],
            'CFOP': infCte.findtext('ns:ide/ns:CFOP', namespaces=ns) or '',
            'Tipo Serviço': infCte.findtext('ns:ide/ns:tpServ', namespaces=ns) or '',
            'Emitente': infCte.findtext('ns:emit/ns:xNome', namespaces=ns) or '',
            'CNPJ Emitente': infCte.findtext('ns:emit/ns:CNPJ', namespaces=ns) or '',
            'UF Remetente': infCte.findtext('ns:rem/ns:enderReme/ns:UF', namespaces=ns) or '',
            'Remetente': infCte.findtext('ns:rem/ns:xNome', namespaces=ns) or '',
            'Destinatário': infCte.findtext('ns:dest/ns:xNome', namespaces=ns) or '',
            'UF Destinatário': infCte.findtext('ns:dest/ns:enderDest/ns:UF', namespaces=ns) or '',
            'Valor Total': infCte.findtext('ns:vPrest/ns:vTPrest', namespaces=ns) or '',
            'Chave de acesso': infCte.get('Id')[3:] if infCte.get('Id') else ''
        }

        return [dados]
    except Exception as e:
        st.error(f"Erro ao processar CTe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

# --- Main Application ---
def main():
    # Page configuration
    st.set_page_config(
        page_title="Conversor XML para Excel", 
        layout="wide",
        page_icon="📄"
    )
    
    # Simple title with your name below
    st.title("Conversor XML para Excel")
    st.caption("Desenvolvido por Patricia Gutierrez")
    
    # Document type selection
    tipo_doc = st.radio(
        "Selecione o tipo de documento:", 
        ["NFe", "CTe"], 
        horizontal=True,
        index=0
    )
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Selecione o arquivo ZIP com os XMLs", 
        type="zip"
    )

    if uploaded_file:
        with st.spinner("Processando arquivos..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded zip
                zip_path = os.path.join(temp_dir, uploaded_file.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Extract XMLs
                xml_files = extrair_xmls_de_zip(zip_path, temp_dir)
                
                if not xml_files:
                    st.warning("Nenhum arquivo XML encontrado no ZIP.")
                else:
                    st.info(f"{len(xml_files)} arquivo(s) XML encontrado(s)")
                    
                    # Process files with progress bar
                    progress_bar = st.progress(0)
                    dados_totais = []
                    
                    for i, xml_file in enumerate(xml_files):
                        progress_bar.progress((i + 1) / len(xml_files))
                        
                        if tipo_doc == "NFe":
                            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                            dados_totais.extend(processar_nfe(xml_file, ns))
                        else:
                            dados_totais.extend(processar_cte(xml_file))
                    
                    # Show results
                    if dados_totais:
                        df = pd.DataFrame(dados_totais)
                        st.dataframe(df)
                        
                        # Create and download Excel
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                            df.to_excel(tmp.name, index=False, engine='openpyxl')
                            with open(tmp.name, 'rb') as f:
                                excel_data = f.read()
                            os.unlink(tmp.name)
                        
                        st.download_button(
                            "Baixar Excel",
                            data=excel_data,
                            file_name=f"{tipo_doc}_Resultado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

if __name__ == "__main__":
    main()
