import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import tempfile
import streamlit as st
from streamlit.runtime.scriptrunner import get_script_run_ctx

def main():
    # Configuração inicial da página
    st.set_page_config(
        page_title="Conversor XML NFe para Excel - Patricia Gutierrez",
        layout="wide",
        page_icon="📄"
    )
    
    # Cabeçalho com créditos
    st.title("📄 Conversor XML NFe para Excel")
    st.markdown("**Desenvolvido por Patricia Gutierrez**")
    st.markdown("Transforme arquivos XML de NFe em planilhas Excel organizadas.")

    # Opções de entrada
    st.sidebar.header("Opções de Entrada")
    input_method = st.sidebar.radio(
        "Selecione o método de entrada:",
        ("Upload de arquivo ZIP", "Selecionar pasta com XMLs")
    )

    # Namespace para parsing XML
    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

    if input_method == "Upload de arquivo ZIP":
        zip_file = st.file_uploader(
            "📂 Selecione o arquivo ZIP com XMLs", 
            type=["zip"],
            help="Arquivo ZIP contendo os XMLs de NFe"
        )
        
        if zip_file:
            with st.spinner('Processando arquivo ZIP...'):
                with tempfile.TemporaryDirectory() as temp_dir:
                    zip_path = os.path.join(temp_dir, zip_file.name)
                    with open(zip_path, "wb") as f:
                        f.write(zip_file.getbuffer())
                    
                    processar_arquivos(zip_path, temp_dir, ns)

    else:  # Selecionar pasta com XMLs
        st.warning("No Streamlit Cloud, a seleção de pastas locais é limitada. "
                 "Recomenda-se usar a opção de upload de ZIP.")
        
        # Alternativa para ambiente local
        if get_script_run_ctx() and not st.runtime.exists():
            xml_dir = st.text_input(
                "Digite o caminho completo da pasta com XMLs:",
                help="Exemplo: C:/pasta/xmls ou /home/usuario/xmls"
            )
            
            if xml_dir and os.path.isdir(xml_dir):
                arquivos_xml = [os.path.join(xml_dir, f) for f in os.listdir(xml_dir) 
                             if f.lower().endswith('.xml')]
                
                if arquivos_xml:
                    with st.spinner(f'Processando {len(arquivos_xml)} XML(s)...'):
                        processar_xmls_diretamente(arquivos_xml, ns)
                else:
                    st.error("Nenhum arquivo XML encontrado na pasta especificada.")
            elif xml_dir:
                st.error("Pasta não encontrada. Verifique o caminho.")

def processar_arquivos(zip_path, temp_dir, ns):
    try:
        arquivos_xml = extrair_xmls_de_zip(zip_path, temp_dir)
        
        if not arquivos_xml:
            st.warning("Nenhum arquivo XML encontrado no ZIP.")
            return
        
        st.success(f"✔️ {len(arquivos_xml)} XML(s) encontrado(s). Processando...")
        
        dados_totais = []
        progress_bar = st.progress(0)
        
        for i, xml_path in enumerate(arquivos_xml):
            progress_bar.progress((i + 1) / len(arquivos_xml))
            dados = processar_xml(xml_path, ns)
            if dados:
                dados_totais.extend(dados)
        
        if dados_totais:
            exibir_e_salvar_resultados(dados_totais)
        else:
            st.warning("Nenhum dado válido encontrado nos XMLs.")
            
    except Exception as e:
        st.error(f"Erro ao processar arquivos: {str(e)}")

def processar_xmls_diretamente(arquivos_xml, ns):
    try:
        dados_totais = []
        progress_bar = st.progress(0)
        
        for i, xml_path in enumerate(arquivos_xml):
            progress_bar.progress((i + 1) / len(arquivos_xml))
            dados = processar_xml(xml_path, ns)
            if dados:
                dados_totais.extend(dados)
        
        if dados_totais:
            exibir_e_salvar_resultados(dados_totais)
        else:
            st.warning("Nenhum dado válido encontrado nos XMLs.")
            
    except Exception as e:
        st.error(f"Erro ao processar arquivos: {str(e)}")

def exibir_e_salvar_resultados(dados_totais):
    df = pd.DataFrame(dados_totais)
    
    # Exibir preview dos dados
    st.subheader("Pré-visualização dos Dados")
    st.dataframe(df.head())
    
    # Opções para salvar
    st.subheader("Salvar Resultados")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Salvar em Excel (download)
        nome_padrao = f"NFe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        nome_arquivo = st.text_input(
            "Nome do arquivo Excel:",
            value=nome_padrao
        )
        
        excel_bytes = criar_excel_bytes(df)
        st.download_button(
            label="⬇️ Baixar Excel",
            data=excel_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        # Opção para salvar localmente (apenas em execução local)
        if get_script_run_ctx() and not st.runtime.exists():
            save_path = st.text_input(
                "Caminho para salvar (opcional):",
                help="Exemplo: C:/pasta/resultados ou /home/usuario/resultados"
            )
            
            if save_path and st.button("💾 Salvar Localmente"):
                try:
                    os.makedirs(save_path, exist_ok=True)
                    full_path = os.path.join(save_path, nome_arquivo)
                    df.to_excel(full_path, index=False)
                    st.success(f"Arquivo salvo em: {full_path}")
                except Exception as e:
                    st.error(f"Erro ao salvar: {str(e)}")

def criar_excel_bytes(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            bytes_data = f.read()
        os.unlink(tmp.name)
    return bytes_data

# ... (mantenha as funções extrair_xmls_de_zip, formatar_valor, obter_status_nfe e processar_xml do código original)

if __name__ == "__main__":
    main()
