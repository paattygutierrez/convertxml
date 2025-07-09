import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Conversor XML para Excel2", layout="wide", page_icon="📄")

st.markdown("""
    <style>
        html, body, [class*="css"] {
            background-color: white !important;
            color: black !important;
        }
    </style>
""", unsafe_allow_html=True)

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

def obter_cnpj_remetente(infNFe, ns):
    transp = infNFe.find('ns:transp', ns)
    if transp is not None:
        remetente = transp.find('ns:rem', ns)
        if remetente is not None:
            return remetente.findtext('ns:CNPJ', default='', namespaces=ns)
    return ''

def obter_cnpj_remetente_por_root(root, ns):
    transp = root.find('.//ns:transp', ns)
    if transp is not None:
        remetente = transp.find('ns:rem', ns)
        if remetente is not None:
            return remetente.findtext('ns:CNPJ', default='', namespaces=ns)
    return ''

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

        cnpj_remetente = obter_cnpj_remetente(infNFe, ns)
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
                'CNPJ Remetente': cnpj_remetente,
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
