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
    # Tentar encontrar o protocolo em diferentes locais
    protNFe = root.find('.//ns:protNFe', ns) or root.find('.//protNFe')
    if protNFe is not None:
        cStat = protNFe.findtext('.//ns:cStat', namespaces=ns) or protNFe.findtext('.//cStat')
        xMotivo = protNFe.findtext('.//ns:xMotivo', namespaces=ns) or protNFe.findtext('.//xMotivo')
        return f"{cStat} - {xMotivo}" if cStat and xMotivo else ''
    return ''

def processar_xml(caminho_xml, ns):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        # Tentar encontrar a NFe em diferentes namespaces
        infNFe = (root.find('.//ns:infNFe', ns) or 
                 root.find('.//{http://www.portalfiscal.inf.br/nfe}infNFe') or
                 root.find('infNFe'))
        
        if infNFe is None:
            st.warning(f"Não foi possível encontrar infNFe no arquivo {os.path.basename(caminho_xml)}")
            return None

        ide = infNFe.find('ns:ide', ns) or infNFe.find('{http://www.portalfiscal.inf.br/nfe}ide') or infNFe.find('ide')
        emit = infNFe.find('ns:emit', ns) or infNFe.find('{http://www.portalfiscal.inf.br/nfe}emit') or infNFe.find('emit')
        itens = infNFe.findall('ns:det', ns) or infNFe.findall('{http://www.portalfiscal.inf.br/nfe}det') or infNFe.findall('det')
        status = obter_status_nfe(root, ns)

        if ide is None or emit is None:
            st.warning(f"Elementos essenciais não encontrados no XML {os.path.basename(caminho_xml)}")
            return None

        dados_por_cfop = {}

        for item in itens:
            prod = item.find('ns:prod', ns) or item.find('{http://www.portalfiscal.inf.br/nfe}prod') or item.find('prod')
            imposto = item.find('ns:imposto', ns) or item.find('{http://www.portalfiscal.inf.br/nfe}imposto') or item.find('imposto')

            if prod is None:
                continue

            cfop = (prod.findtext('ns:CFOP', default='', namespaces=ns) or
                   prod.findtext('{http://www.portalfiscal.inf.br/nfe}CFOP', default='') or
                   prod.findtext('CFOP', default=''))
            
            vProd = (prod.findtext('ns:vProd', default='0', namespaces=ns) or
                    prod.findtext('{http://www.portalfiscal.inf.br/nfe}vProd', default='0') or
                    prod.findtext('vProd', default='0'))
            
            vICMS = '0'
            pICMS = ''
            if imposto is not None:
                vICMS = (imposto.findtext('.//ns:vICMS', default='0', namespaces=ns) or
                        imposto.findtext('.//{http://www.portalfiscal.inf.br/nfe}vICMS', default='0') or
                        imposto.findtext('.//vICMS', default='0'))
                
                pICMS = (imposto.findtext('.//ns:pICMS', default='', namespaces=ns) or
                        imposto.findtext('.//{http://www.portalfiscal.inf.br/nfe}pICMS', default='') or
                        imposto.findtext('.//pICMS', default=''))

            pIPI = (imposto.findtext('.//ns:IPI/ns:IPITrib/ns:pIPI', default='', namespaces=ns) if imposto is not None else '')
            pPIS = (imposto.findtext('.//ns:PIS/ns:PISAliq/ns:pPIS', default='', namespaces=ns) if imposto is not None else '')
            pCOFINS = (imposto.findtext('.//ns:COFINS/ns:COFINSAliq/ns:pCOFINS', default='', namespaces=ns) if imposto is not None else '')

            if cfop not in dados_por_cfop:
                nNF = (ide.findtext('ns:nNF', default='', namespaces=ns) or
                      ide.findtext('{http://www.portalfiscal.inf.br/nfe}nNF', default='') or
                      ide.findtext('nNF', default=''))
                
                dhEmi = (ide.findtext('ns:dhEmi', default='', namespaces=ns) or
                        ide.findtext('{http://www.portalfiscal.inf.br/nfe}dhEmi', default='') or
                        ide.findtext('dhEmi', default=''))
                
                cnpj = (emit.findtext('ns:CNPJ', default='', namespaces=ns) or
                        emit.findtext('{http://www.portalfiscal.inf.br/nfe}CNPJ', default='') or
                        emit.findtext('CNPJ', default=''))
                
                xNome = (emit.findtext('ns:xNome', default='', namespaces=ns) or
                        emit.findtext('{http://www.portalfiscal.inf.br/nfe}xNome', default='') or
                        emit.findtext('xNome', default=''))
                
                uf = (emit.findtext('ns:enderEmit/ns:UF', default='', namespaces=ns) or
                      emit.findtext('{http://www.portalfiscal.inf.br/nfe}enderEmit/{http://www.portalfiscal.inf.br/nfe}UF', default='') or
                      emit.findtext('enderEmit/UF', default=''))

                dados_por_cfop[cfop] = {
                    'Numero NFe': nNF,
                    'Data Emissão': dhEmi[:10] if dhEmi else '',
                    'CNPJ Emitente': cnpj,
                    'Emitente': xNome,
                    'UF Emitente': uf,
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

        total = (infNFe.find('ns:total/ns:ICMSTot', ns) or 
                infNFe.find('{http://www.portalfiscal.inf.br/nfe}total/{http://www.portalfiscal.inf.br/nfe}ICMSTot') or
                infNFe.find('total/ICMSTot'))
        
        if total is not None:
            for cfop_dado in dados_por_cfop.values():
                cfop_dado['Vlr IPI'] = formatar_valor(total.findtext('ns:vIPI', default='0', namespaces=ns) or total.findtext('vIPI', default='0')
                cfop_dado['Vlr PIS'] = formatar_valor(total.findtext('ns:vPIS', default='0', namespaces=ns) or total.findtext('vPIS', default='0')
                cfop_dado['Vlr COFINS'] = formatar_valor(total.findtext('ns:vCOFINS', default='0', namespaces=ns) or total.findtext('vCOFINS', default='0')
                cfop_dado['VICMS ST'] = formatar_valor(total.findtext('ns:vST', default='0', namespaces=ns) or total.findtext('vST', default='0')
                cfop_dado['Vlr Frete'] = formatar_valor(total.findtext('ns:vFrete', default='0', namespaces=ns) or total.findtext('vFrete', default='0'))
                cfop_dado['Vlr Seguro'] = formatar_valor(total.findtext('ns:vSeg', default='0', namespaces=ns) or total.findtext('vSeg', default='0'))

                cfop_dado['Vlr Nota'] = formatar_valor(str(cfop_dado['Vlr Nota']))
                cfop_dado['Vlr ICMS'] = formatar_valor(str(cfop_dado['Vlr ICMS']))

        return list(dados_por_cfop.values())

    except Exception as e:
        st.error(f"Erro ao processar {os.path.basename(caminho_xml)}: {str(e)}")
        return None

# -------- Streamlit App --------

def main():
    st.set_page_config(page_title="Conversor XML NFe para Excel", layout="wide")
    st.title("📄 XML NFe para Excel")
    st.markdown("Faça o upload de um arquivo `.zip` contendo **arquivos XML de NFe** para gerar um Excel com os dados consolidados.")

    zip_file = st.file_uploader("📂 Selecione o arquivo ZIP com XMLs", type=["zip"])

    if zip_file is not None:
        with st.spinner('Processando arquivos...'):
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_path = os.path.join(temp_dir, zip_file.name)
                with open(zip_path, "wb") as f:
                    f.write(zip_file.getbuffer())

                try:
                    arquivos_xml = extrair_xmls_de_zip(zip_path, temp_dir)
                    if not arquivos_xml:
                        st.warning("Nenhum arquivo XML encontrado no ZIP.")
                        return
                    
                    st.success(f"✔️ {len(arquivos_xml)} XML(s) encontrado(s) no arquivo ZIP.")
                    
                    # Namespace padrão
                    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                    
                    dados_totais = []
                    progress_bar = st.progress(0)
                    for i, xml_path in enumerate(arquivos_xml):
                        progress_bar.progress((i + 1) / len(arquivos_xml))
                        dados = processar_xml(xml_path, ns)
                        if dados:
                            dados_totais.extend(dados)

                    if dados_totais:
                        df = pd.DataFrame(dados_totais)
                        st.dataframe(df.head())  # Mostra apenas as primeiras linhas para preview

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
                        st.warning("Nenhuma NFe válida encontrada nos arquivos XML.")
                except Exception as e:
                    st.error(f"Ocorreu um erro ao processar o arquivo: {str(e)}")
                    st.error("Verifique se os arquivos XML são válidos e no formato correto.")

if __name__ == "__main__":
    main()
