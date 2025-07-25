import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# Função para extrair os arquivos XML dos ZIPs
def extrair_xmls_de_zips(pasta):
    temp_dir = tempfile.mkdtemp()
    for raiz, _, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            if arquivo.lower().endswith('.zip'):
                caminho_zip = os.path.join(raiz, arquivo)
                with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
    return temp_dir

# Formatação padrão
def formatar_valor(valor):
    return valor.replace('.', ',') if valor else '0,00'

# Processar dados por item
def processar_nfe_por_item(caminho_xml):
    tree = ET.parse(caminho_xml)
    root = tree.getroot()
    ns = {'ns': root.tag.split('}')[0].strip('{')}

    infNFe = root.find('.//ns:infNFe', ns)
    if infNFe is None:
        return []

    ide = infNFe.find('ns:ide', ns)
    emit = infNFe.find('ns:emit', ns)
    dest = infNFe.find('ns:dest', ns)

    chave = infNFe.attrib.get('Id', '')[-44:]
    numero_nf = ide.findtext('ns:nNF', default='', namespaces=ns)
    serie = ide.findtext('ns:serie', default='', namespaces=ns)
    data_emissao = ide.findtext('ns:dhEmi', default='', namespaces=ns)
    if not data_emissao:
        data_emissao = ide.findtext('ns:dEmi', default='', namespaces=ns)
    data_emissao = data_emissao[:10]

    emitente = emit.findtext('ns:xNome', default='', namespaces=ns)
    cnpj_emit = emit.findtext('ns:CNPJ', default='', namespaces=ns)
    uf_emit = emit.findtext('ns:enderEmit/ns:UF', default='', namespaces=ns)

    cnpj_dest = dest.findtext('ns:CNPJ', default='', namespaces=ns)

    inf_adic = infNFe.find('ns:infAdic', ns)
    obs_nf = inf_adic.findtext('ns:infCpl', default='', namespaces=ns) if inf_adic is not None else ''

    itens = infNFe.findall('ns:det', ns)
    dados = []
    for item in itens:
        prod = item.find('ns:prod', ns)
        imposto = item.find('ns:imposto', ns)

        cod_prod = prod.findtext('ns:cProd', default='', namespaces=ns)
        desc = prod.findtext('ns:xProd', default='', namespaces=ns)
        ncm = prod.findtext('ns:NCM', default='', namespaces=ns)
        cfop = prod.findtext('ns:CFOP', default='', namespaces=ns)
        unidade = prod.findtext('ns:uCom', default='', namespaces=ns)
        qtd = prod.findtext('ns:qCom', default='', namespaces=ns)
        vlr_unit = prod.findtext('ns:vUnCom', default='', namespaces=ns)
        vlr_total = prod.findtext('ns:vProd', default='', namespaces=ns)
        vlr_desc = prod.findtext('ns:vDesc', default='', namespaces=ns)
        vlr_frete = prod.findtext('ns:vFrete', default='', namespaces=ns)
        vlr_seguro = prod.findtext('ns:vSeg', default='', namespaces=ns)
        obs_item = prod.findtext('ns:infAdProd', default='', namespaces=ns)
        cBenef = prod.findtext('ns:cBenef', default='', namespaces=ns)

        vBC = pICMS = vICMS = vBCST = pST = vST = vPIS = vCOFINS = vICMSDeson = '0'

        if imposto is not None:
            icms = imposto.find('ns:ICMS', ns)
            if icms is not None:
                for tipo_icms in icms:
                    vBC = tipo_icms.findtext('ns:vBC', default=vBC, namespaces=ns)
                    pICMS = tipo_icms.findtext('ns:pICMS', default=pICMS, namespaces=ns)
                    vICMS = tipo_icms.findtext('ns:vICMS', default=vICMS, namespaces=ns)
                    vICMSDeson = tipo_icms.findtext('ns:vICMSDeson', default='0', namespaces=ns)
                    vBCST = tipo_icms.findtext('ns:vBCST', default=vBCST, namespaces=ns)
                    pST = tipo_icms.findtext('ns:pICMSST', default=pST, namespaces=ns)
                    vST = tipo_icms.findtext('ns:vICMSST', default=vST, namespaces=ns)

            pis = imposto.find('ns:PIS/ns:PISAliq', ns)
            if pis is not None:
                vPIS = pis.findtext('ns:vPIS', default=vPIS, namespaces=ns)

            cofins = imposto.find('ns:COFINS/ns:COFINSAliq', ns)
            if cofins is not None:
                vCOFINS = cofins.findtext('ns:vCOFINS', default=vCOFINS, namespaces=ns)

        dados.append({
            'Chave': chave,
            'Numero NF': numero_nf,
            'Serie': serie,
            'Data': data_emissao,
            'Emitente': emitente,
            'CNPJ Emitente': cnpj_emit,
            'CNPJ Destinatário': cnpj_dest,
            'CFOP': cfop,
            'Codigo Produto': cod_prod,
            'Desc': desc,
            'NCM': ncm,
            'Obs Item': obs_item,
            'Qtd': qtd,
            'unidade': unidade,
            'Vlr Unit': vlr_unit,
            'Vlr total': vlr_total,
            'Base ICMS': vBC,
            'Aliquota': pICMS,
            'Vlr ICMS': vICMS,
            'Base ICMS ST': vBCST,
            'Vlr ICMS ST': vST,
            'Vlr PIS': vPIS,
            'Vlr COFINS': vCOFINS,
            'Vlr Frete': vlr_frete,
            'Vlr Seguro': vlr_seguro,
            'Vlr Desconto': vlr_desc,
            'Obs NFe': obs_nf,
            'cBenef': cBenef,
            'ICMS Desonerado': formatar_valor(vICMSDeson)
        })
    return dados

# Processar por cabeçalho da NFe
def processar_nfe_por_cabecalho(caminho_xml):
    tree = ET.parse(caminho_xml)
    root = tree.getroot()
    ns = {'ns': root.tag.split('}')[0].strip('{')}

    infNFe = root.find('.//ns:infNFe', ns)
    if infNFe is None:
        return []

    ide = infNFe.find('ns:ide', ns)
    emit = infNFe.find('ns:emit', ns)
    total = infNFe.find('ns:total/ns:ICMSTot', ns)

    chave = infNFe.attrib.get('Id', '')[-44:]
    numero_nf = ide.findtext('ns:nNF', default='', namespaces=ns)
    serie = ide.findtext('ns:serie', default='', namespaces=ns)
    data_emissao = ide.findtext('ns:dhEmi', default='', namespaces=ns)
    if not data_emissao:
        data_emissao = ide.findtext('ns:dEmi', default='', namespaces=ns)
    data_emissao = data_emissao[:10]

    emitente = emit.findtext('ns:xNome', default='', namespaces=ns)
    cnpj_emit = emit.findtext('ns:CNPJ', default='', namespaces=ns)
    uf_emit = emit.findtext('ns:enderEmit/ns:UF', default='', namespaces=ns)

    valor_total = total.findtext('ns:vNF', default='0', namespaces=ns)
    icms = total.findtext('ns:vICMS', default='0', namespaces=ns)
    frete = total.findtext('ns:vFrete', default='0', namespaces=ns)
    seguro = total.findtext('ns:vSeg', default='0', namespaces=ns)

    return [{
        'Chave': chave,
        'Numero NF': numero_nf,
        'Serie': serie,
        'Data': data_emissao,
        'Emitente': emitente,
        'CNPJ Emitente': cnpj_emit,
        'UF Emitente': uf_emit,
        'Valor Total': valor_total,
        'ICMS': icms,
        'Frete': frete,
        'Seguro': seguro
    }]

# Função principal para gerar Excel
def gerar_excel(tipo):
    pasta_origem = filedialog.askdirectory(title="Selecione a pasta com os ZIPs")
    if not pasta_origem:
        return

    temp_dir = extrair_xmls_de_zips(pasta_origem)
    dados_gerais = []

    for raiz, _, arquivos in os.walk(temp_dir):
        for arquivo in arquivos:
            if arquivo.endswith('.xml'):
                caminho_xml = os.path.join(raiz, arquivo)
                try:
                    if tipo == 'Item':
                        dados = processar_nfe_por_item(caminho_xml)
                    else:
                        dados = processar_nfe_por_cabecalho(caminho_xml)
                    dados_gerais.extend(dados)
                except Exception as e:
                    print(f"Erro ao processar {arquivo}: {e}")

    if dados_gerais:
        df = pd.DataFrame(dados_gerais)
        nome_arquivo = f"NFe_{tipo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho_salvar = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=nome_arquivo,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if caminho_salvar:
            df.to_excel(caminho_salvar, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho_salvar}")
    else:
        messagebox.showwarning("Aviso", "Nenhum dado encontrado nos arquivos.")

# Interface gráfica com Tkinter
janela = tk.Tk()
janela.title("Conversor de XML NFe")

tk.Label(janela, text="Selecione o tipo de extração:").pack(pady=10)
tk.Button(janela, text="NFe por Item", width=30, command=lambda: gerar_excel('Item')).pack(pady=5)
tk.Button(janela, text="NFe por Cabeçalho", width=30, command=lambda: gerar_excel('Cabecalho')).pack(pady=5)
tk.Button(janela, text="Sair", width=30, command=janela.destroy).pack(pady=20)

janela.mainloop()
