import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Conversor XML para Excel", layout="wide", page_icon="📄")

def extrair_xmls_de_zip(zip_path, destino):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(destino)
    return [os.path.join(raiz, nome)
            for raiz, _, arquivos in os.walk(destino)
            for nome in arquivos if nome.lower().endswith('.xml')]

def formatar_valor(val):
    if val is None or val == '':
        return '0,00'
    return str(val).replace('.', ',')

def processar_nfe_por_item(caminho_xml, ns):
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
        infAdic = infNFe.find('ns:infAdic', ns)
        transp = infNFe.find('ns:transp', ns)

        remetente = transp.find('ns:rem', ns) if transp is not None else None
        cnpj_remetente = remetente.findtext('ns:CNPJ', default='', namespaces=ns) if remetente is not None else ''
        obs_nfe = infAdic.findtext('ns:infCpl', default='', namespaces=ns) if infAdic is not None else ''

        dados_itens = []
        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)
            infAdicProd = item.find('ns:infAdProd', ns)

            vBC = pICMS = vICMS = vBCST = pST = vST = vPIS = vCOFINS = vFrete = vSeg = vDesc = '0'

            if imposto is not None:
                icms = imposto.find('.//ns:ICMS', ns)
                pis = imposto.find('.//ns:PIS', ns)
                cofins = imposto.find('.//ns:COFINS', ns)
                if icms is not None:
                    for child in icms:
                        if child.tag.endswith('ICMS00') or child.tag.endswith('ICMS20'):
                            vBC = child.findtext('ns:vBC', default='0', namespaces=ns)
                            pICMS = child.findtext('ns:pICMS', default='0', namespaces=ns)
                            vICMS = child.findtext('ns:vICMS', default='0', namespaces=ns)
                        vBCST_item = child.findtext('ns:vBCST', default='0', namespaces=ns)
                        pST_item = child.findtext('ns:pST', default='0', namespaces=ns)
                        vST_item = child.findtext('ns:vST', default='0', namespaces=ns)
                        vBCST = vBCST_item if vBCST_item else vBCST
                        pST = pST_item if pST_item else pST
                        vST = vST_item if vST_item else vST
                if pis is not None:
                    pis_item = pis.find('.//ns:PISAliq', ns) or pis.find('.//ns:PISOutr', ns)
                    if pis_item is not None:
                        vPIS = pis_item.findtext('ns:vPIS', default='0', namespaces=ns)
                if cofins is not None:
                    cofins_item = cofins.find('.//ns:COFINSAliq', ns) or cofins.find('.//ns:COFINSOutr', ns)
                    if cofins_item is not None:
                        vCOFINS = cofins_item.findtext('ns:vCOFINS', default='0', namespaces=ns)

            vFrete_item = prod.findtext('ns:vFrete', default='0', namespaces=ns)
            vFrete = vFrete_item if vFrete_item != '0' else total.findtext('ns:vFrete', default='0', namespaces=ns)
            vSeg_item = prod.findtext('ns:vSeg', default='0', namespaces=ns)
            vSeg = vSeg_item if vSeg_item != '0' else total.findtext('ns:vSeg', default='0', namespaces=ns)
            vDesc_item = prod.findtext('ns:vDesc', default='0', namespaces=ns)
            vDesc = vDesc_item if vDesc_item != '0' else total.findtext('ns:vDesc', default='0', namespaces=ns)

            dados = {
                'Chave': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                'Numero NF': ide.findtext('ns:nNF', default='', namespaces=ns),
                'Serie': ide.findtext('ns:serie', default='', namespaces=ns),
                'Data': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
                'Emitente': emit.findtext('ns:xNome', default='', namespaces=ns),
                'CNPJ Emitente': emit.findtext('ns:CNPJ', default='', namespaces=ns),
                'Remetente CNPJ': cnpj_remetente,
                'CFOP': prod.findtext('ns:CFOP', default='', namespaces=ns),
                'Codigo Produto': prod.findtext('ns:cProd', default='', namespaces=ns),
                'Desc': prod.findtext('ns:xProd', default='', namespaces=ns),
                'NCM': prod.findtext('ns:NCM', default='', namespaces=ns),
                'Obs Item': infAdicProd.text if infAdicProd is not None else '',
                'Qtd': formatar_valor(prod.findtext('ns:qCom', default='0', namespaces=ns)),
                'unidade': prod.findtext('ns:uCom', default='', namespaces=ns),
                'Vlr Unit': formatar_valor(prod.findtext('ns:vUnCom', default='0', namespaces=ns)),
                'Vlr total': formatar_valor(prod.findtext('ns:vProd', default='0', namespaces=ns)),
                'Base ICMS': formatar_valor(vBC),
                'Aliquota': formatar_valor(pICMS),
                'Vlr ICMS': formatar_valor(vICMS),
                'Base ICMS ST': formatar_valor(vBCST),
                'Vlr ICMS ST': formatar_valor(vST),
                'Vlr PIS': formatar_valor(vPIS),
                'Vlr COFINS': formatar_valor(vCOFINS),
                'Vlr Frete': formatar_valor(vFrete),
                'Vlr Seguro': formatar_valor(vSeg),
                'Vlr Desconto': formatar_valor(vDesc),
                'Obs NFe': obs_nfe
            }
            dados_itens.append(dados)

        return dados_itens

    except Exception as e:
        st.error(f"Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

def criar_excel(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            data = f.read()
        os.unlink(tmp.name)
    return data

def main():
    st.title("Conversor XML para Excel")
    st.caption("Desenvolvido por Patricia Gutierrez")

    tipo_doc = st.radio("Selecione o tipo de documento:", ["NFe", "CTe"], horizontal=True)
    layout = ""
    if tipo_doc == "NFe":
        layout = st.radio("Layout de saída:", ["Cabeçalho", "Por Item"], horizontal=True)

    uploaded_file = st.file_uploader("Selecione o arquivo ZIP com os XMLs", type="zip")

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
                                continue  # ainda não adaptado
                            else:
                                dados_totais.extend(processar_nfe_por_item(xml_file, ns))
                        else:
                            continue  # ainda não adaptado para CTe

                    if dados_totais:
                        colunas_ordenadas = [
                            'Chave', 'Numero NF', 'Serie', 'Data', 'Emitente', 'CNPJ Emitente', 'Remetente CNPJ',
                            'CFOP', 'Codigo Produto', 'Desc', 'NCM', 'Obs Item', 'Qtd', 'unidade', 'Vlr Unit',
                            'Vlr total', 'Base ICMS', 'Aliquota', 'Vlr ICMS', 'Base ICMS ST', 'Vlr ICMS ST',
                            'Vlr PIS', 'Vlr COFINS', 'Vlr Frete', 'Vlr Seguro', 'Vlr Desconto', 'Obs NFe']

                        df = pd.DataFrame(dados_totais)
                        df = df.reindex(columns=colunas_ordenadas).fillna('')
                        st.dataframe(df)
                        excel_data = criar_excel(df)
                        nome_arquivo = f"{tipo_doc}_{layout.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        st.download_button("Baixar Excel", data=excel_data, file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
