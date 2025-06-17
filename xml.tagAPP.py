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
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# --- CSS Personalizado ---
st.markdown("""
    <style>
        /* Barra de progresso personalizada */
        .stProgress > div > div > div > div {
            height: 15px;
            background-color: #4CAF50;
            border-radius: 10px;
        }
        
        /* Estilo dos títulos */
        .app-title {
            display: flex;
            align-items: center;
            gap: 15px;
            color: #2e7d32;
            border-bottom: 2px solid #4CAF50;
            padding-bottom: 10px;
        }
        
        .excel-icon {
            font-size: 32px;
            color: #217346;
        }
        
        /* Rodapé */
        .footer {
            text-align: center;
            padding: 15px;
            background-color: #f0f2f6;
            font-size: 12px;
            margin-top: 30px;
            border-radius: 5px;
        }
        
        .dev-name {
            color: #217346;
            font-weight: bold;
            font-size: 1.1em;
        }
    </style>
""", unsafe_allow_html=True)

# --- Funções (mantenha as mesmas funções do código anterior) ---
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

# ... (Mantenha todas as outras funções existentes: processar_nfe, processar_cte, etc.)

# --- Interface Principal ---
def main():
    # Título com ícone de Excel
    st.markdown("""
    <div class="app-title">
        <span class="excel-icon">📊</span>
        <h1>Conversor XML para Excel</h1>
    </div>
    """, unsafe_allow_html=True)
    
    # Barra de autoria destacada
    st.markdown("""
    <div style="background-color:#e8f5e9;padding:10px;border-radius:5px;margin-bottom:20px;">
        <h3 style="color:#217346;text-align:center;">Desenvolvido por <span class="dev-name">PATRICIA GUTIERREZ</span></h3>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<p style="font-size:16px;">Converta seus arquivos XML de <b>NFe</b> ou <b>CTe</b> para planilhas Excel</p>', unsafe_allow_html=True)
    
    with st.expander("ℹ️ Instruções de Uso", expanded=False):
        st.write("""
        1. Selecione o tipo de documento (NFe ou CTe)
        2. Faça upload do arquivo ZIP contendo os XMLs
        3. Aguarde o processamento automático
        4. Baixe o arquivo Excel gerado
        """)
    
    # Seção de upload e processamento
    col1, col2 = st.columns([1, 3])
    with col1:
        tipo_doc = st.radio("Tipo de documento:", ["NFe", "CTe"], horizontal=True)
    
    uploaded_file = st.file_uploader(
        "Selecione o arquivo ZIP com os XMLs", 
        type="zip",
        help="Arquivo compactado contendo os XMLs a serem processados"
    )

    if uploaded_file:
        with st.spinner("Preparando para processar..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_path = os.path.join(temp_dir, uploaded_file.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                xml_files = extrair_xmls_de_zip(zip_path, temp_dir)

                if not xml_files:
                    st.warning("⚠️ Nenhum arquivo XML encontrado no ZIP.")
                else:
                    st.success(f"🔍 {len(xml_files)} arquivo(s) XML encontrado(s)")
                    
                    # Barra de progresso
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    dados_totais = []
                    for i, xml in enumerate(xml_files):
                        progress = (i + 1) / len(xml_files)
                        progress_bar.progress(progress)
                        status_text.text(f"📂 Processando arquivo {i+1} de {len(xml_files)}...")
                        
                        if tipo_doc == "NFe":
                            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                            dados_totais.extend(processar_nfe(xml, ns))
                        else:
                            dados_totais.extend(processar_cte(xml))
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    if dados_totais:
                        df = pd.DataFrame(dados_totais)
                        
                        st.subheader("📊 Resultados Processados")
                        st.dataframe(df.head())
                        
                        st.subheader("📥 Download")
                        excel_data = criar_excel(df)
                        st.download_button(
                            label="⬇️ Baixar Planilha Excel",
                            data=excel_data,
                            file_name=f"{tipo_doc}_Resultado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Clique para baixar o arquivo Excel com os dados processados"
                        )

    # Rodapé
    st.markdown("""
    <div class="footer">
        Sistema desenvolvido por <span class="dev-name">PATRICIA GUTIERREZ</span> | Versão 2.0 | © 2023
    </div>
    """, unsafe_allow_html=True)

def criar_excel(df):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        with open(tmp.name, 'rb') as f:
            data = f.read()
        os.unlink(tmp.name)
    return data

if __name__ == "__main__":
    main()
