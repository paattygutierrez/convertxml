import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from datetime import datetime

# --- Configuração da Página ---
st.set_page_config(page_title="Conversor XML - Excel", layout="wide", page_icon="📄")

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

# ---------- NFe ----------
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
        st.error(f"❌ Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

# ---------- CTe ----------
def format_number(value):
    if value and any(c.isdigit() for c in value):
        return value.replace('.', ',')
    return value

def get_text(element, path, ns):
    tag = element.find(path, ns)
    return format_number(tag.text) if tag is not None and tag.text else None

def get_valor_icms(element, ns):
    icms_total = 0.0
    icms_tags = [
        'ICMS00', 'ICMS10', 'ICMS20', 'ICMS30', 'ICMS45', 'ICMS60', 
        'ICMS70', 'ICMS90', 'ICMSPart', 'ICMSST', 'ICMSUFFim', 
        'ICMSUFRemet', 'ICMSOutraUF'
    ]
    campos_icms = ['vICMS', 'vICMSST', 'vICMSUFFim', 'vICMSUFRemet', 'vICMSOutraUF']
    for tag in icms_tags:
        for campo in campos_icms:
            valor = element.find(f'ns:imp/ns:ICMS/ns:{tag}/ns:{campo}', ns)
            if valor is not None and valor.text:
                try:
                    icms_total += float(valor.text)
                except ValueError:
                    pass
    return format_number(str(icms_total)) if icms_total > 0 else None

def processar_cte(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}

        infCte = root.find('.//ns:infCte', ns)
        if infCte is None:
            return []

        dados = {
            'Número CTe': get_text(infCte, 'ns:ide/ns:nCT', ns),
            'Data Emissão': get_text(infCte, 'ns:ide/ns:dhEmi', ns)[:10],
            'CFOP': get_text(infCte, 'ns:ide/ns:CFOP', ns),
            'Tipo Serviço': get_text(infCte, 'ns:ide/ns:tpServ', ns),
            'Emitente': get_text(infCte, 'ns:emit/ns:xNome', ns),
            'CNPJ Emitente': get_text(infCte, 'ns:emit/ns:CNPJ', ns),
            'UF Remetente': get_text(infCte, 'ns:rem/ns:enderReme/ns:UF', ns),
            'Remetente': get_text(infCte, 'ns:rem/ns:xNome', ns),
            'Destinatário': get_text(infCte, 'ns:dest/ns:xNome', ns),
            'UF Destinatário': get_text(infCte, 'ns:dest/ns:enderDest/ns:UF', ns),
            'Valor Total': get_text(infCte, 'ns:vPrest/ns:vTPrest', ns),
            'Valor ICMS': get_valor_icms(infCte, ns),
            'Chave de acesso': infCte.get('Id')[3:] if infCte.get('Id') else ''
        }

        return [dados]
    except Exception as e:
        st.error(f"❌ Erro ao processar CTe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

# --- Interface ---
st.title("📄 Conversor XML - Excel")
st.markdown('<p style="font-size:14px; color:gray">Converta seus arquivos XML de <b>NFe</b> ou <b>CTe</b> para Excel. <br><b>Desenvolvido por Patricia Gutierrez</b></p>', unsafe_allow_html=True)

tipo_doc = st.selectbox("Selecione o tipo de documento:", ["NFe", "CTe"])
uploaded_file = st.file_uploader("Selecione o arquivo ZIP com os XMLs", type="zip")

def criar_excel(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            data = f.read()
        os.unlink(tmp.name)
    return data

if uploaded_file:
    with st.spinner("Processando..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, uploaded_file.name)
            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            xml_files = extrair_xmls_de_zip(zip_path, temp_dir)

            if not xml_files:
                st.warning("Nenhum XML encontrado.")
            else:
                dados_totais = []
                for i, xml in enumerate(xml_files):
                    st.progress((i + 1) / len(xml_files))
                    if tipo_doc == "NFe":
                        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                        dados_totais.extend(processar_nfe(xml, ns))
                    else:
                        dados_totais.extend(processar_cte(xml))

                if dados_totais:
                    df = pd.DataFrame(dados_totais)
                    st.dataframe(df)

                    excel_data = criar_excel(df)
                    st.download_button("⬇️ Baixar Excel", data=excel_data,
                        file_name=f"{tipo_doc}_XMLs.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
