import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import tempfile
import streamlit as st

def extrair_xmls_de_zip(zip_path, destino):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(destino)

    arquivos_xml = []
    for raiz, _, arquivos in os.walk(destino):
        for nome in arquivos:
            if nome.lower().endswith('.xml'):
                arquivos_xml.append(os.path.join(raiz, nome))
    return arquivos_xml

def formatar_valor(val):
    if val is None:
        return '0,00'
    return val.replace('.', ',') if '.' in val else val

def obter_status_nfe(root, ns):
    protNFe = root.find('.//ns:protNFe', ns)
    if protNFe is not None:
        cStat = protNFe.findtext('.//ns:cStat', namespaces=ns)
        xMotivo = protNFe.findtext('.//ns:xMotivo', namespaces=ns)
        return f"{cStat} - {xMotivo}" if cStat and xMotivo else ''
    return ''

def processar_xml(caminho_xml, ns):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        infNFe = root.find('.//ns:infNFe', ns)
        if infNFe is None:
            return None

        ide = infNFe.find('ns:ide', ns)
        emit = infNFe.find('ns:emit', ns)
        itens = infNFe.findall('ns:det', ns)
        status = obter_status_nfe(root, ns)

        dados_por_cfop = {}

        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)

            cfop = prod.findtext('ns:CFOP', default='', namespaces=ns)
            vProd = prod.findtext('ns:vProd', default='0', namespaces=ns)
            vICMS = imposto.findtext('.//ns:vICMS', default='0', namespaces=ns) if imposto is not None else '0'
            pICMS = imposto.findtext('.//ns:pICMS', default='', namespaces=ns)

            pIPI = imposto.findtext('.//ns:IPI/ns:IPITrib/ns:pIPI', default='', namespaces=ns)
            pPIS = imposto.findtext('.//ns:PIS/ns:PISAliq/ns:pPIS', default='', namespaces=ns)
            pCOFINS = imposto.findtext('.//ns:COFINS/ns:COFINSAliq/ns:pCOFINS', default='', namespaces=ns)

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
                    'Aliquota IPI': pIPI,
                    'Vlr PIS': 0.0,
                    'Aliquota PIS': pPIS,
                    'Vlr COFINS': 0.0,
                    'Aliquota COFINS': pCOFINS,
                    'VICMS ST': 0.0,
                    'Vlr Frete': 0.0,
                    'Vlr Seguro': 0.0,
                    'Chave de acesso': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                    'Status': status
                }

            dados_por_cfop[cfop]['Vlr Nota'] += float(vProd.replace(',', '.'))
            dados_por_cfop[cfop]['Vlr ICMS'] += float(vICMS.replace(',', '.'))

        total = infNFe.find('ns:total/ns:ICMSTot', ns)
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
        st.error(f"Erro ao processar {os.path.basename(caminho_xml)}: {str(e)}")
        return None

# -------- Streamlit App --------

st.set_page_config(page_title="Conversor XML NFe para Excel - Patricia Gutierrez", layout="wide")
st.title("📄 XML NFe para Excel")
st.markdown("Faça o upload de um arquivo `.zip` contendo **arquivos XML de NFe** para gerar um Excel com os dados consolidados.")

zip_file = st.file_uploader("📂 Selecione o arquivo ZIP com XMLs", type=["zip"])

if zip_file:
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, zip_file.name)
        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        try:
            arquivos_xml = extrair_xmls_de_zip(zip_path, temp_dir)
            st.success(f"✔️ {len(arquivos_xml)} XMLs encontrados no arquivo ZIP.")
            
            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
            dados_totais = []
            for xml_path in arquivos_xml:
                dados = processar_xml(xml_path, ns)
                if dados:
                    dados_totais.extend(dados)

            if dados_totais:
                df = pd.DataFrame(dados_totais)
                st.dataframe(df)

                nome_arquivo = f"NFe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                excel_path = os.path.join(temp_dir, nome_arquivo)
                df.to_excel(excel_path, index=False)

                with open(excel_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=f,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.warning("Nenhuma NFe válida encontrada.")
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
