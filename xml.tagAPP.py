import streamlit as st
import pandas as pd
import zipfile
import os
import xml.etree.ElementTree as ET
import tempfile
from io import BytesIO
from datetime import datetime

def extrair_xmls_de_zip(zip_path, extract_path):
    """
    Extrai todos os arquivos XML de um arquivo ZIP para um diretório temporário.
    """
    xml_files = []
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
    for root, _, files in os.walk(extract_path):
        for file in files:
            if file.endswith('.xml'):
                xml_files.append(os.path.join(root, file))
    return xml_files

def processar_nfe_por_item(xml_path, ns):
    """
    Processa um arquivo XML de NFe extraindo dados por item.
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        emit = root.find('.//ns:emit', ns)
        dest = root.find('.//ns:dest', ns)  # Captura o destinatário
        entrega = root.find('.//ns:entrega', ns)  # Captura o local de entrega (se houver)
        ide = root.find('.//ns:ide', ns)
        total = root.find('.//ns:total', ns)
        det_list = root.findall('.//ns:det', ns)

        if emit is None or ide is None or total is None:
            return []

        chave_acesso_tag = root.find('.//ns:infProt/ns:chNFe', ns)
        chave_acesso = chave_acesso_tag.text if chave_acesso_tag is not None else ""

        status_tag = root.find('.//ns:infProt/ns:cStat', ns)
        status = status_tag.text if status_tag is not None else ""

        emitente = emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else ""
        cnpj_emitente = emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else ""
        uf_emitente = emit.find('ns:enderEmit/ns:UF', ns).text if emit.find('ns:enderEmit/ns:UF', ns) is not None else ""
        
        # Extrai a UF do Destinatário
        uf_destino = ""
        if dest is not None and dest.find('ns:enderDest/ns:UF', ns) is not None:
            uf_destino = dest.find('ns:enderDest/ns:UF', ns).text

        # Extrai a UF de Entrega (se informada)
        uf_entrega = ""
        if entrega is not None and entrega.find('ns:UF', ns) is not None:
            uf_entrega = entrega.find('ns:UF', ns).text

        numero_nfe = ide.find('ns:nNF', ns).text if ide.find('ns:nNF', ns) is not None else ""
        data_emissao = ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else ""

        infadic = root.find('.//ns:infAdic', ns)
        observacoes = ""
        if infadic is not None:
            inf_cpl = infadic.find('ns:infCpl', ns)
            inf_fisco = infadic.find('ns:infAdFisco', ns)
            observacoes = (inf_cpl.text if inf_cpl is not None else "")
            if inf_fisco is not None:
                observacoes += (" | " + inf_fisco.text if inf_cpl is not None else inf_fisco.text)

        dados = []
        for det in det_list:
            prod = det.find('ns:prod', ns)
            imposto = det.find('ns:imposto', ns)

            if prod is None or imposto is None:
                continue

            # Código do Produto/Item
            codigo_item = prod.find('ns:cProd', ns)
            codigo_item_val = codigo_item.text if codigo_item is not None else ""

            icms = imposto.find('.//ns:ICMS', ns)
            icms_valor = icms.find('.//ns:vICMS', ns) if icms is not None else None
            icms_aliquota = icms.find('.//ns:pICMS', ns) if icms is not None else None
            icms_cst = icms.find('.//ns:CST', ns) if icms is not None else None
            icms_desonerado = icms.find('.//ns:vICMSDeson', ns) if icms is not None else None

            ipi_valor = imposto.find('.//ns:IPI/ns:IPITrib/ns:vIPI', ns)
            pis_valor = imposto.find('.//ns:PIS/ns:PISAliq/ns:vPIS', ns)
            cofins_valor = imposto.find('.//ns:COFINS/ns:COFINSAliq/ns:vCOFINS', ns)
            icms_st_valor = imposto.find('.//ns:ICMS/*/ns:vICMSST', ns)

            cbenef = prod.find('ns:cBenef', ns)
            cfop = prod.find('ns:CFOP', ns)

            frete = root.find('.//ns:transp/ns:vFrete', ns)
            seguro = root.find('.//ns:transp/ns:vSeg', ns)

            # IBS / CBS 
            ibs_valor = imposto.find('.//ns:vIBS', ns)
            ibs_cst = imposto.find('.//ns:CST', ns)
            cbs_valor = imposto.find('.//ns:vCBS', ns)
            cbs_cst = imposto.find('.//ns:CST', ns)

            dados.append({
                "Número NFe": numero_nfe,
                "Data de Emissão": data_emissao,
                "CNPJ Emitente": cnpj_emitente,
                "Emitente": emitente,
                "UF Emitente": uf_emitente,
                "UF Destino": uf_destino,
                "UF Entrega": uf_entrega,
                "Código do Item": codigo_item_val,  # <--- NOVO CAMPO ACRESCENTADO AQUI
                "Valor da Nota": total.find('ns:ICMSTot/ns:vNF', ns).text if total.find('ns:ICMSTot/ns:vNF', ns) is not None else "",
                "ICMS": icms_valor.text if icms_valor is not None else "",
                "Alíquota ICMS": icms_aliquota.text if icms_aliquota is not None else "",
                "CST ICMS": icms_cst.text if icms_cst is not None else "",
                "IPI": ipi_valor.text if ipi_valor is not None else "",
                "PIS": pis_valor.text if pis_valor is not None else "",
                "COFINS": cofins_valor.text if cofins_valor is not None else "",
                "ICMS ST": icms_st_valor.text if icms_st_valor is not None else "",
                "IBS": ibs_valor.text if ibs_valor is not None else "",
                "CST IBS": ibs_cst.text if ibs_cst is not None else "",
                "CBS": cbs_valor.text if cbs_valor is not None else "",
                "CST CBS": cbs_cst.text if cbs_cst is not None else "",
                "Frete": frete.text if frete is not None else "",
                "Seguro": seguro.text if seguro is not None else "",
                "Chave de Acesso": chave_acesso,
                "cBenef": cbenef.text if cbenef is not None else "",
                "ICMS Desonerado": icms_desonerado.text if icms_desonerado is not None else "",
                "CFOP": cfop.text if cfop is not None else "",
                "Status da NFe": status,
                "Observações da Nota": observacoes
            })
        return dados
    except Exception as e:
        st.error(f"Erro ao analisar o arquivo XML: {os.path.basename(xml_path)}. Erro: {e}")
        return []

def processar_nfe_por_cabecalho(xml_path, ns):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        emit = root.find('.//ns:emit', ns)
        ide = root.find('.//ns:ide', ns)
        total = root.find('.//ns:total', ns)

        if emit is None or ide is None or total is None:
            return []

        chave_acesso_tag = root.find('.//ns:infProt/ns:chNFe', ns)
        chave_acesso = chave_acesso_tag.text if chave_acesso_tag is not None else ""
        status_tag = root.find('.//ns:infProt/ns:cStat', ns)
        status = status_tag.text if status_tag is not None else ""
        emitente = emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else ""
        cnpj_emitente = emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else ""
        uf_emitente = emit.find('ns:enderEmit/ns:UF', ns).text if emit.find('ns:enderEmit/ns:UF', ns) is not None else ""
        numero_nfe = ide.find('ns:nNF', ns).text if ide.find('ns:nNF', ns) is not None else ""
        data_emissao = ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else ""

        frete = root.find('.//ns:transp/ns:vFrete', ns)
        seguro = root.find('.//ns:transp/ns:vSeg', ns)

        ibs_total = total.find('ns:IBSCBSTot/ns:vIBS', ns)
        cbs_total = total.find('ns:IBSCBSTot/ns:vCBS', ns)

        return [{
            "Número NFe": numero_nfe,
            "Data de Emissão": data_emissao,
            "CNPJ Emitente": cnpj_emitente,
            "Emitente": emitente,
            "UF Emitente": uf_emitente,
            "Valor da Nota": total.find('ns:ICMSTot/ns:vNF', ns).text if total.find('ns:ICMSTot/ns:vNF', ns) is not None else "",
            "ICMS": total.find('ns:ICMSTot/ns:vICMS', ns).text if total.find('ns:ICMSTot/ns:vICMS', ns) is not None else "",
            "Alíquota ICMS": "",
            "CST ICMS": "",
            "IPI": total.find('ns:ICMSTot/ns:vIPI', ns).text if total.find('ns:ICMSTot/ns:vIPI', ns) is not None else "",
            "PIS": total.find('ns:ICMSTot/ns:vPIS', ns).text if total.find('ns:ICMSTot/ns:vPIS', ns) is not None else "",
            "COFINS": total.find('ns:ICMSTot/ns:vCOFINS', ns).text if total.find('ns:ICMSTot/ns:vCOFINS', ns) is not None else "",
            "ICMS ST": total.find('ns:ICMSTot/ns:vST', ns).text if total.find('ns:ICMSTot/ns:vST', ns) is not None else "",
            "IBS": ibs_total.text if ibs_total is not None else "",
            "CBS": cbs_total.text if cbs_total is not None else "",
            "Frete": frete.text if frete is not None else "",
            "Seguro": seguro.text if seguro is not None else "",
            "Chave de Acesso": chave_acesso,
            "Status da NFe": status
        }]
    except:
        return []

def processar_cte(xml_path, ns):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        ide = root.find('.//ns:ide', ns)
        emit = root.find('.//ns:emit', ns)
        valor_total = root.find('.//ns:vTPrest', ns)
        icms = root.find('.//ns:ICMS00', ns)
        chave_acesso_tag = root.find('.//ns:infProt/ns:chCTe', ns)

        if ide is None or emit is None or valor_total is None:
            return []

        return [{
            "Número CTe": ide.find('ns:nCT', ns).text if ide.find('ns:nCT', ns) is not None else "",
            "Data de Emissão": ide.find('ns:dhEmi', ns).text if ide.find('ns:dhEmi', ns) is not None else "",
            "CNPJ Emitente": emit.find('ns:CNPJ', ns).text if emit.find('ns:CNPJ', ns) is not None else "",
            "Emitente": emit.find('ns:xNome', ns).text if emit.find('ns:xNome', ns) is not None else "",
            "Valor Total": valor_total.text if valor_total is not None else "",
            "Chave de Acesso": chave_acesso_tag.text if chave_acesso_tag is not None else ""
        }]
    except:
        return []

def main():
    st.title("XML to EXCEL")

    tipo_doc = st.radio("Tipo de Documento:", ["NFe", "CTe"])
    layout = st.radio("Layout de Exportação:", ["Item", "Cabeçalho"])

    uploaded_files = st.file_uploader(
        "Selecione um ou mais arquivos ZIP com os XMLs",
        type="zip",
        accept_multiple_files=True
    )

    if uploaded_files:
        with st.spinner("Processando arquivos..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                xml_files = []
                for uploaded_file in uploaded_files:
                    zip_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(zip_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    arquivos_extraidos = extrair_xmls_de_zip(zip_path, temp_dir)
                    xml_files.extend(arquivos_extraidos)

                if not xml_files:
                    st.warning("Nenhum arquivo XML encontrado.")
                else:
                    dados_totais = []
                    progress_bar = st.progress(0)
                    for i, xml_file in enumerate(xml_files):
                        progress_bar.progress((i + 1) / len(xml_files))
                        if tipo_doc == "NFe":
                            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                            if layout == "Cabeçalho":
                                dados_totais.extend(processar_nfe_por_cabecalho(xml_file, ns))
                            else:
                                dados_totais.extend(processar_nfe_por_item(xml_file, ns))
                        else:
                            ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}
                            dados_totais.extend(processar_cte(xml_file, ns))

                    if dados_totais:
                        df = pd.DataFrame(dados_totais)
                        if 'Data de Emissão' in df.columns:
                            df['Data de Emissão'] = pd.to_datetime(df['Data de Emissão'], errors='coerce', utc=True).dt.date
                        
                        st.subheader("Dados Extraídos")
                        st.dataframe(df)

                        csv = df.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("Baixar CSV", csv, "resultado.csv", "text/csv")

    st.markdown("---")
    st.markdown("Desenvolvido por Patricia Gutierrez")

if __name__ == "__main__":
    main()
