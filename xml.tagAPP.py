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
    """
    Extrai todos os arquivos XML de um arquivo ZIP para um diretório de destino.
    Retorna uma lista dos caminhos completos dos arquivos XML extraídos.
    """
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(destino)
    return [os.path.join(raiz, nome)
            for raiz, _, arquivos in os.walk(destino)
            for nome in arquivos if nome.lower().endswith('.xml')]

def formatar_valor(val):
    """
    Formata um valor numérico para o padrão brasileiro (vírgula como separador decimal).
    Se o valor for None ou vazio, retorna '0,00'.
    """
    if val is None or val == '':
        return '0,00'
    # Converte para float primeiro para garantir que a formatação seja consistente
    try:
        val_float = float(str(val).replace(',', '.'))
        return f"{val_float:.2f}".replace('.', ',')
    except ValueError:
        return '0,00' # Retorna 0,00 se a conversão falhar

def processar_nfe_por_item(caminho_xml, ns):
    """
    Processa um arquivo XML de NFe, extraindo dados detalhados por item.
    Inclui cBenef e ICMS Desonerado por item.
    """
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        infNFe = root.find('.//ns:infNFe', ns)
        if infNFe is None:
            return []

        ide = infNFe.find('ns:ide', ns)
        emit = infNFe.find('ns:emit', ns)
        dest = infNFe.find('ns:dest', ns) or infNFe.find('ns:destinatario', ns)
        cnpj_destinatario = dest.findtext('ns:CNPJ', default='', namespaces=ns) if dest is not None else ''
        itens = infNFe.findall('ns:det', ns)
        total = infNFe.find('ns:total/ns:ICMSTot', ns)
        infAdic = infNFe.find('ns:infAdic', ns)

        obs_nfe = infAdic.findtext('ns:infCpl', default='', namespaces=ns) if infAdic is not None else ''

        dados_itens = []
        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)
            infAdicProd = item.find('ns:infAdProd', ns)
            
            # Inicializa todos os campos com valores padrão
            vBC = '0'
            pICMS = '0'
            vICMS = '0'
            vBCST = '0'
            pST = '0'
            vST = '0'
            vPIS = '0'
            vCOFINS = '0'
            vFrete = '0'
            vSeg = '0'
            vDesc = '0'
            
            vICMSDeson = '0' # Nova coluna ICMS Desonerado

            # cBenef é um filho direto de 'prod'
            cBenef = prod.findtext('ns:cBenef', default='', namespaces=ns) # Correção aqui

            # Processa impostos
            if imposto is not None:
                icms = imposto.find('ns:ICMS', ns)
                if icms is not None:
                    # Verifica todos os tipos possíveis de ICMS
                    for tipo_icms in icms:
                        # ICMS Normal
                        vBC_current = tipo_icms.findtext('ns:vBC', default=vBC, namespaces=ns)
                        pICMS_current = tipo_icms.findtext('ns:pICMS', default=pICMS, namespaces=ns)
                        vICMS_current = tipo_icms.findtext('ns:vICMS', default=vICMS, namespaces=ns)
                        
                        # Atualiza apenas se o valor não for padrão '0' (para pegar o mais relevante)
                        if vBC_current != '0': vBC = vBC_current
                        if pICMS_current != '0': pICMS = pICMS_current
                        if vICMS_current != '0': vICMS = vICMS_current
                        
                        # ICMS ST (pode estar em vários tipos de ICMS)
                        vBCST_current = tipo_icms.findtext('ns:vBCST', default=vBCST, namespaces=ns)
                        pST_current = tipo_icms.findtext('ns:pICMSST', default=pST, namespaces=ns)
                        vST_current = tipo_icms.findtext('ns:vICMSST', default=vST, namespaces=ns)

                        if vBCST_current != '0': vBCST = vBCST_current
                        if pST_current != '0': pST = pST_current
                        if vST_current != '0': vST = vST_current

                        # ICMS Desonerado
                        vICMSDeson_current = tipo_icms.findtext('ns:vICMSDeson', default='0', namespaces=ns)
                        if vICMSDeson_current != '0': vICMSDeson = vICMSDeson_current

                # PIS e COFINS
                pis = imposto.find('ns:PIS', ns)
                if pis is not None:
                    pis_item = pis.find('.//ns:PISAliq', ns) or pis.find('.//ns:PISOutr', ns) or pis.find('.//ns:PISNT', ns)
                    if pis_item is not None:
                        vPIS = pis_item.findtext('ns:vPIS', default='0', namespaces=ns)

                cofins = imposto.find('ns:COFINS', ns)
                if cofins is not None:
                    cofins_item = cofins.find('.//ns:COFINSAliq', ns) or cofins.find('.//ns:COFINSOutr', ns) or cofins.find('.//ns:COFINSNT', ns)
                    if cofins_item is not None:
                        vCOFINS = cofins_item.findtext('ns:vCOFINS', default='0', namespaces=ns)

            # Se não encontrou valores de ST no item, tenta pegar do total (caso raro, mas para robustez)
            if vBCST == '0':
                vBCST = total.findtext('ns:vBCST', default='0', namespaces=ns)
            if vST == '0':
                vST = total.findtext('ns:vST', default='0', namespaces=ns)

            # Frete, Seguro e Desconto (pega do item ou do total, priorizando o item)
            vFrete_item = prod.findtext('ns:vFrete', default='0', namespaces=ns)
            vFrete = vFrete_item if vFrete_item != '0' else total.findtext('ns:vFrete', default='0', namespaces=ns)
            
            vSeg_item = prod.findtext('ns:vSeg', default='0', namespaces=ns)
            vSeg = vSeg_item if vSeg_item != '0' else total.findtext('ns:vSeg', default='0', namespaces=ns)
            
            vDesc_item = prod.findtext('ns:vDesc', default='0', namespaces=ns)
            vDesc = vDesc_item if vDesc_item != '0' else total.findtext('ns:vDesc', default='0', namespaces=ns)

            # Dados do item
            dados = {
                'Chave': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                'Numero NF': ide.findtext('ns:nNF', default='', namespaces=ns),
                'Serie': ide.findtext('ns:serie', default='', namespaces=ns),
                'Data': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
                'Emitente': emit.findtext('ns:xNome', default='', namespaces=ns),
                'CNPJ Emitente': emit.findtext('ns:CNPJ', default='', namespaces=ns),
                'CNPJ Destinatário': cnpj_destinatario,
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
                'cBenef': cBenef, # Nova coluna
                'ICMS Desonerado': formatar_valor(vICMSDeson), # Nova coluna
                'Obs NFe': obs_nfe
            }

            dados_itens.append(dados)

        return dados_itens

    except Exception as e:
        st.error(f"Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

def processar_nfe_por_cabecalho(caminho_xml, ns):
    """
    Processa um arquivo XML de NFe, agregando dados por CFOP (cabeçalho da nota).
    ICMS Desonerado é somado por CFOP; cBenef não é aplicável e fica em branco.
    """
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
        dest = infNFe.find('ns:dest', ns) or infNFe.find('ns:destinatario', ns)
        cnpj_destinatario = dest.findtext('ns:CNPJ', default='', namespaces=ns) if dest is not None else ''

        dados_por_cfop = {}
        for item in itens:
            prod = item.find('ns:prod', ns)
            imposto = item.find('ns:imposto', ns)
            icms = imposto.find('ns:ICMS', ns) if imposto is not None else None

            cfop = prod.findtext('ns:CFOP', default='', namespaces=ns)
            
            # Inicializa valores para o item atual a serem somados
            vBC_item_current = 0.0
            pICMS_item_current = '0' # Alíquota é a última encontrada para o item
            vICMS_item_current = 0.0
            vBCST_item_current = 0.0
            pST_item_current = '0' # Alíquota ST é a última encontrada para o item
            vST_item_current = 0.0
            vICMSDeson_item_current = 0.0 # Novo: Valor do ICMS Desonerado para o item

            if icms is not None:
                for tipo_icms in icms:
                    # ICMS normal
                    vBC_val = tipo_icms.findtext('ns:vBC', default='0', namespaces=ns)
                    pICMS_val = tipo_icms.findtext('ns:pICMS', namespaces=ns)
                    vICMS_val = tipo_icms.findtext('ns:vICMS', default='0', namespaces=ns)
                    
                    if vBC_val:
                        vBC_item_current += float(vBC_val.replace(',', '.'))
                    if pICMS_val:
                        pICMS_item_current = pICMS_val
                    if vICMS_val:
                        vICMS_item_current += float(vICMS_val.replace(',', '.'))
                    
                    # ICMS ST
                    vBCST_val = tipo_icms.findtext('ns:vBCST', namespaces=ns)
                    pST_val = tipo_icms.findtext('ns:pICMSST', namespaces=ns)
                    vST_val = tipo_icms.findtext('ns:vICMSST', namespaces=ns)
                    
                    if vBCST_val:
                        vBCST_item_current += float(vBCST_val.replace(',', '.'))
                    if pST_val:
                        pST_item_current = pST_val
                    if vST_val:
                        vST_item_current += float(vST_val.replace(',', '.'))

                    # ICMS Desonerado - NOVO
                    vICMSDeson_val = tipo_icms.findtext('ns:vICMSDeson', default='0', namespaces=ns)
                    if vICMSDeson_val:
                        vICMSDeson_item_current += float(vICMSDeson_val.replace(',', '.'))


            if cfop not in dados_por_cfop:
                dados_por_cfop[cfop] = {
                    'Chave': infNFe.get('Id')[3:] if infNFe.get('Id') else '',
                    'Numero NF': ide.findtext('ns:nNF', default='', namespaces=ns),
                    'Serie': ide.findtext('ns:serie', default='', namespaces=ns),
                    'Data': ide.findtext('ns:dhEmi', default='', namespaces=ns)[:10],
                    'Emitente': emit.findtext('ns:xNome', default='', namespaces=ns),
                    'CNPJ Emitente': emit.findtext('ns:CNPJ', default='', namespaces=ns),
                    'CNPJ Destinatário': cnpj_destinatario,
                    'CFOP': cfop,
                    'Codigo Produto': '',  # Não aplicável no cabeçalho
                    'Desc': '',  # Não aplicável no cabeçalho
                    'NCM': '',  # Não aplicável no cabeçalho
                    'Obs Item': '',  # Não aplicável no cabeçalho
                    'Qtd': '0,00',  # Não aplicável no cabeçalho
                    'unidade': '',  # Não aplicável no cabeçalho
                    'Vlr Unit': '0,00',  # Não aplicável no cabeçalho
                    'Vlr total': '0,00',  # Será preenchido depois
                    'Base ICMS': '0,00',  # Será preenchido depois
                    'Aliquota': pICMS_item_current, # Use alíquota do item (última encontrada)
                    'Vlr ICMS': '0,00',  # Será preenchido depois
                    'Base ICMS ST': '0,00',  # Será preenchido depois
                    'Vlr ICMS ST': '0,00',  # Será preenchido depois
                    'Vlr PIS': formatar_valor(total.findtext('ns:vPIS', default='0', namespaces=ns)),
                    'Vlr COFINS': formatar_valor(total.findtext('ns:vCOFINS', default='0', namespaces=ns)),
                    'Vlr Frete': formatar_valor(total.findtext('ns:vFrete', default='0', namespaces=ns)),
                    'Vlr Seguro': formatar_valor(total.findtext('ns:vSeg', default='0', namespaces=ns)),
                    'Vlr Desconto': formatar_valor(total.findtext('ns:vDesc', default='0', namespaces=ns)),
                    'cBenef': '', # Não aplicável para resumo de cabeçalho
                    'ICMS Desonerado': '0,00', # Novo, será somado
                    'Obs NFe': infAdic.findtext('ns:infCpl', default='', namespaces=ns) if infAdic is not None else ''
                }

            # Atualiza valores acumulados para o CFOP
            dados_por_cfop[cfop]['Vlr total'] = str(float(dados_por_cfop[cfop]['Vlr total'].replace(',', '.')) + float(prod.findtext('ns:vProd', default='0', namespaces=ns).replace(',', '.')))
            dados_por_cfop[cfop]['Base ICMS'] = str(float(dados_por_cfop[cfop]['Base ICMS'].replace(',', '.')) + vBC_item_current)
            dados_por_cfop[cfop]['Vlr ICMS'] = str(float(dados_por_cfop[cfop]['Vlr ICMS'].replace(',', '.')) + vICMS_item_current)
            dados_por_cfop[cfop]['Base ICMS ST'] = str(float(dados_por_cfop[cfop]['Base ICMS ST'].replace(',', '.')) + vBCST_item_current)
            dados_por_cfop[cfop]['Vlr ICMS ST'] = str(float(dados_por_cfop[cfop]['Vlr ICMS ST'].replace(',', '.')) + vST_item_current)
            dados_por_cfop[cfop]['ICMS Desonerado'] = str(float(dados_por_cfop[cfop]['ICMS Desonerado'].replace(',', '.')) + vICMSDeson_item_current) # Novo

        # Formata os valores acumulados
        for cfop_dado in dados_por_cfop.values():
            cfop_dado['Vlr total'] = formatar_valor(cfop_dado['Vlr total'])
            cfop_dado['Base ICMS'] = formatar_valor(cfop_dado['Base ICMS'])
            cfop_dado['Vlr ICMS'] = formatar_valor(cfop_dado['Vlr ICMS'])
            cfop_dado['Base ICMS ST'] = formatar_valor(cfop_dado['Base ICMS ST'])
            cfop_dado['Vlr ICMS ST'] = formatar_valor(cfop_dado['Vlr ICMS ST'])
            cfop_dado['Aliquota'] = formatar_valor(cfop_dado['Aliquota']) # A alíquota é do último item processado para aquele CFOP, não uma soma.
            cfop_dado['ICMS Desonerado'] = formatar_valor(cfop_dado['ICMS Desonerado']) # Novo

        return list(dados_por_cfop.values())

    except Exception as e:
        st.error(f"Erro ao processar NFe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

def processar_cte(caminho_xml, ns):
    """
    Processa um arquivo XML de CTe, extraindo dados relevantes.
    As colunas cBenef e ICMS Desonerado não são aplicáveis para CTe.
    """
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        infCte = root.find('.//ns:infCte', ns)
        if infCte is None:
            return []

        # Find the dest (Destinatário) tag
        dest = infCte.find('ns:dest', ns)
        cnpj_destinatario = dest.findtext('ns:CNPJ', namespaces=ns) if dest is not None else ''

        dados = {
            'Chave': infCte.get('Id')[3:] if infCte.get('Id') else '',
            'Numero NF': infCte.findtext('ns:ide/ns:nCT', namespaces=ns) or '',
            'Serie': infCte.findtext('ns:ide/ns:serie', namespaces=ns) or '',
            'Data': (infCte.findtext('ns:ide/ns:dhEmi', namespaces=ns) or '')[:10],
            'Emitente': infCte.findtext('ns:emit/ns:xNome', namespaces=ns) or '',
            'CNPJ Emitente': infCte.findtext('ns:emit/ns:CNPJ', namespaces=ns) or '',
            'CNPJ Destinatário': cnpj_destinatario,
            'CFOP': infCte.findtext('ns:ide/ns:CFOP', namespaces=ns) or '',
            'Codigo Produto': '',
            'Desc': '',
            'NCM': '',
            'Obs Item': '',
            'Qtd': '',
            'unidade': '',
            'Vlr Unit': '',
            'Vlr total': formatar_valor(infCte.findtext('ns:vPrest/ns:vTPrest', namespaces=ns) or '0'),
            'Base ICMS': '',
            'Aliquota': '',
            'Vlr ICMS': '',
            'Base ICMS ST': '',
            'Vlr ICMS ST': '',
            'Vlr PIS': '',
            'Vlr COFINS': '',
            'Vlr Frete': '',
            'Vlr Seguro': '',
            'Vlr Desconto': '',
            'cBenef': '', # Não aplicável para CTe
            'ICMS Desonerado': '', # Não aplicável para CTe
            'Obs NFe': ''
        }

        return [dados]
    except Exception as e:
        st.error(f"Erro ao processar CTe {os.path.basename(caminho_xml)}: {str(e)}")
        return []

def criar_excel(df):
    """
    Cria um arquivo Excel a partir de um DataFrame Pandas em memória.
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            data = f.read()
        os.unlink(tmp.name)
    return data

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
    layout = ""
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
                            else: # Layout "Por Item"
                                dados_totais.extend(processar_nfe_por_item(xml_file, ns))
                        else: # Tipo de documento "CTe"
                            ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}
                            dados_totais.extend(processar_cte(xml_file, ns))
                    
                    if dados_totais:
                        # Definir a ordem das colunas, incluindo as novas
                        colunas_ordenadas = [
                            'Chave', 'Numero NF', 'Serie', 'Data', 'Emitente', 'CNPJ Emitente', 'CNPJ Destinatário',
                            'CFOP', 'Codigo Produto', 'Desc', 'NCM', 'Obs Item', 'Qtd', 'unidade',
                            'Vlr Unit', 'Vlr total', 'Base ICMS', 'Aliquota', 'Vlr ICMS',
                            'Base ICMS ST', 'Vlr ICMS ST', 'Vlr PIS', 'Vlr COFINS', 'Vlr Frete',
                            'Vlr Seguro', 'Vlr Desconto',
                            'cBenef', # Nova coluna
                            'ICMS Desonerado', # Nova coluna
                            'Obs NFe'
                        ]
                        
                        df = pd.DataFrame(dados_totais)
                        
                        # Reordenar as colunas e preencher valores faltantes
                        # Use .reindex para garantir que todas as colunas estejam presentes na ordem definida
                        # Mesmo que uma coluna não tenha sido preenchida para um tipo de doc (ex: cBenef no CTe)
                        df = df.reindex(columns=colunas_ordenadas).fillna('')
                        
                        # Aplicar filtros apenas para NFe
                        if tipo_doc == "NFe":
                            st.subheader("Filtros")
                            
                            # Converter a coluna Data para datetime, ignorando erros
                            df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
                            
                            # Criar colunas para os filtros
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # Filtro de CFOP
                                cfops_disponiveis = sorted(df['CFOP'].unique())
                                cfops_selecionados = st.multiselect(
                                    "Selecione CFOP(s)",
                                    options=cfops_disponiveis,
                                    default=cfops_disponiveis
                                )
                            
                            with col2:
                                # Filtro de data
                                datas_validas = df[pd.notna(df['Data'])]
                                if not datas_validas.empty:
                                    data_min = datas_validas['Data'].min()
                                    data_max = datas_validas['Data'].max()
                                    
                                    datas_selecionadas = st.date_input(
                                        "Selecione o intervalo de datas",
                                        value=[data_min.to_pydatetime().date(), data_max.to_pydatetime().date()],
                                        min_value=data_min.to_pydatetime().date(),
                                        max_value=data_max.to_pydatetime().date()
                                    )
                                    
                                    data_inicio, data_fim = None, None
                                    if len(datas_selecionadas) == 2:
                                        data_inicio = pd.to_datetime(datas_selecionadas[0])
                                        data_fim = pd.to_datetime(datas_selecionadas[1])
                                else:
                                    st.warning("Nenhuma data válida encontrada nos XMLs")
                            
                            # Aplicar filtros
                            df_filtrado = df.copy()
                            
                            # Filtro de CFOP
                            if cfops_selecionados:
                                df_filtrado = df_filtrado[df_filtrado['CFOP'].isin(cfops_selecionados)]
                            
                            # Filtro de data (apenas se as datas forem válidas)
                            if data_inicio and data_fim:
                                df_filtrado = df_filtrado[
                                    (df_filtrado['Data'] >= data_inicio) & 
                                    (df_filtrado['Data'] <= data_fim)
                                ]
                            
                            # Mostrar contagem de registros
                            st.info(f"Total de registros: {len(df_filtrado)}")
                            st.dataframe(df_filtrado)
                            df_para_excel = df_filtrado
                        else:
                            st.dataframe(df)
                            df_para_excel = df
                        
                        excel_data = criar_excel(df_para_excel)
                        
                        # Nome do arquivo condicional
                        if tipo_doc == "NFe":
                            nome_arquivo = f"{tipo_doc}_{layout.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        else:
                            nome_arquivo = f"{tipo_doc}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        
                        st.download_button(
                            "Baixar Excel",
                            data=excel_data,
                            file_name=nome_arquivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("Nenhum dado foi processado a partir dos XMLs. Verifique se os arquivos estão corretos.")

if __name__ == "__main__":
    main()

