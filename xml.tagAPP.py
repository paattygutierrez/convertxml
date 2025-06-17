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

# --- Funções Auxiliares ---
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

# --- Processamento NFe Cabeçalho ---
def processar_nfe_por_cabecalho(caminho_xml, ns):
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
                    'Chave NFe': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
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
                    'Observacoes': infAdic.findtext('ns:infCpl', default='', namespaces=ns) if infAdic is not None else ''
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

# --- Processamento NFe Por Item ---
def processar_nfe_por_item(caminho_xml, ns):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        infNFe = root.find('.//ns:infNFe', ns)
        if infNFe is None:
            return []

        ide = infNFe.find('ns:ide', ns)
        emit = infNFe.find('ns:emit', ns)
        dest = infNFe.find('ns:dest', ns) or infNFe.find('ns:destinatario', ns)
        itens = infNFe.findall('ns:det', ns)
        total = infNFe.find('ns:total/ns:ICMSTot', ns)
        infAdic = infNFe.find('ns:infAdic', ns)
        transp = infNFe.find('ns:transp', ns)
        cobr = infNFe.find('ns:cobr', ns)

        dados_itens = []
        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)
            infAdicProd = item.find('ns:infAdProd', ns)
            
            # Dados básicos do item
            dados = {
                'Chave NFe': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                'Número NFe': ide.findtext('ns:nNF', default='', namespaces=ns),
                'Data Emissão': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
                'Modelo': ide.findtext('ns:mod', default='', namespaces=ns),
                'Série': ide.findtext('ns:serie', default='', namespaces=ns),
                'Tipo Operação': ide.findtext('ns:tpNF', default='', namespaces=ns),
                'Emitente': emit.findtext('ns:xNome', default='', namespaces=ns),
                'CNPJ Emitente': emit.findtext('ns:CNPJ', default='', namespaces=ns),
                'Destinatário': dest.findtext('ns:xNome', default='', namespaces=ns) if dest is not None else '',
                'CNPJ Destinatário': dest.findtext('ns:CNPJ', default='', namespaces=ns) if dest is not None else '',
                'CFOP': prod.findtext('ns:CFOP', default='', namespaces=ns),
                'Código Produto': prod.findtext('ns:cProd', default='', namespaces=ns),
                'Descrição': prod.findtext('ns:xProd', default='', namespaces=ns),
                'NCM': prod.findtext('ns:NCM', default='', namespaces=ns),
                'CEST': prod.findtext('ns:CEST', default='', namespaces=ns),
                'Quantidade': prod.findtext('ns:qCom', default='', namespaces=ns),
                'Unidade': prod.findtext('ns:uCom', default='', namespaces=ns),
                'Valor Unitário': formatar_valor(prod.findtext('ns:vUnCom', default='0', namespaces=ns)),
                'Valor Total': formatar_valor(prod.findtext('ns:vProd', default='0', namespaces=ns)),
                'Observações Item': infAdicProd.text if infAdicProd is not None else '',
                'Observações NFe': infAdic.findtext('ns:infCpl', default='', namespaces=ns) if infAdic is not None else '',
                'Transportador': transp.findtext('ns:transporta/ns:xNome', default='', namespaces=ns) if transp is not None else '',
                'Volume': transp.findtext('ns:vol/ns:qVol', default='', namespaces=ns) if transp is not None else ''
            }

            # Informações de impostos
            if imposto is not None:
                icms = imposto.find('.//ns:ICMS', ns)
                ipi = imposto.find('.//ns:IPI', ns)
                pis = imposto.find('.//ns:PIS', ns)
                cofins = imposto.find('.//ns:COFINS', ns)

                if icms is not None:
                    for child in icms:
                        if child.tag.endswith('ICMS00') or child.tag.endswith('ICMS20'):
                            dados['Alíquota ICMS'] = child.findtext('ns:pICMS', default='', namespaces=ns)
                            dados['Valor ICMS'] = formatar_valor(child.findtext('ns:vICMS', default='0', namespaces=ns))
                            break

                dados['Valor IPI'] = formatar_valor(ipi.findtext('.//ns:vIPI', default='0', namespaces=ns)) if ipi is not None else '0,00'
                dados['Valor PIS'] = formatar_valor(pis.findtext('.//ns:vPIS', default='0', namespaces=ns)) if pis is not None else '0,00'
                dados['Valor COFINS'] = formatar_valor(cofins.findtext('.//ns:vCOFINS', default='0', namespaces=ns)) if cofins is not None else '0,00'

            dados_itens.append(dados)

        return dados_itens

    except Exception as e:
        st.error(f"Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

# --- Processamento CTe ---
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
            'Remetente': infCte.findtext('ns:rem/ns:xNome', namespaces=ns) or '',
            'Destinatário': infCte.findtext('ns:dest/ns:xNome', namespaces=ns) or '',
            'Valor Total': infCte.findtext('ns:vPrest/ns:vTPrest', namespaces=ns) or '',
            'Chave de acesso': infCte.get('Id')[3:] if infCte.get('Id') else ''
        }

        return [dados]
    except Exception as e:
        st.error(f"Erro ao processar CTe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

# --- Interface Principal ---
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
                        nome_arquivo = f"{tipo_doc}_{layout.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        
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
