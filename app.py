import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(page_title="RPA - Gerador de Planilha de Nota", layout="centered")

st.title("📦 RPA Conversor de XML Copacker")
st.write("Atualize a planilha base e converta seus XMLs!")

# ==============================================================================
# ⚙️ CONFIGURAÇÃO DOS SELETORES (Sem espaços após a vírgula para evitar erros no Excel)
# ==============================================================================
opcoes_remetente = '"LONDRINA,ARAMA,POSITIVE CO,SANTA LUZIA"'
opcoes_movimentacao = '"REMESSA PARA INDUSTRIALIZAÇÃO,REMESSA PARA ARMAZENAGEM"'
opcoes_dest_descricao = '"AMENDOAS DO BRASIL LTDA,BEBIDAS POTY LTDA,CASTROLANDA COOPERATIVA AGROIND,CIA IGUAÇU DE CAFE SOLÚVEL,FRYSK INDUSTRIAL LTDA EM RECUPERA,REVPACK TECN COM COMP PLASTICOS LTD,USINA DE LACTICINIOS JUSSARA,EMERGENTCOLD,EXPRESSO MANIR LTDA,GRAN PAR LOGISTICA LTDA,SUPERFRIO ARMAZENS GERAIS SA,EF SOLUÇÕES,MMC INDUSTRIA DE PRODUTOS NUTRATEC"'
# ==============================================================================

# --- FUNÇÃO PARA LIMPEZA EXTREMA DE TEXTO ---
def limpar_texto_comparacao(texto):
    if pd.isna(texto):
        return ""
    txt = str(texto).upper().strip()
    txt = re.sub(r'[.\-\/\,\_\:\(\)]', '', txt)
    txt = txt.replace("LTDA", "").replace("S/A", "").replace("SA", "").replace("S.A", "")
    return " ".join(txt.split())

# --- 1. SEÇÃO DE UPLOAD DA PLANILHA BASE ---
st.header("1️⃣ Atualizar Planilha Base (Trânsito)")
arquivo_base_upload = st.file_uploader("Arraste aqui a planilha 'base_transito.xlsx' atualizada", type=["xlsx"])

df_base = None

if arquivo_base_upload is not None:
    try:
        df_base = pd.read_excel(arquivo_base_upload, sheet_name=0)
        df_base.columns = [str(col).strip().upper() for col in df_base.columns]
        
        col_nf = next((c for c in df_base.columns if "NF" in c), None)
        col_emit = next((c for c in df_base.columns if "EMIT" in c or "FORN" in c), None)
        col_mat = next((c for c in df_base.columns if "MAT" in c), None)
        col_desc_base = "XPROD" if "XPROD" in df_base.columns else next((c for c in df_base.columns if "DESC" in c or "PROD" in c), None)
        
        if col_nf and col_emit and col_mat:
            df_base[col_nf] = df_base[col_nf].apply(lambda x: str(int(float(x))) if re.match(r'^\d+\.\d+$', str(x)) else str(x).strip())
            
            st.success("✅ Planilha Base carregada e ativa na memória do site!")
            
            colunas_exibicao = [col_nf, col_emit, col_mat]
            if col_desc_base and col_desc_base not in colunas_exibicao:
                colunas_exibicao.append(col_desc_base)
                
            df_preview = df_base[colunas_exibicao].tail(5)
            st.dataframe(df_preview, use_container_width=True)
            
            st.session_state['cols_base'] = {'nf': col_nf, 'emit': col_emit, 'mat': col_mat}
        else:
            st.error("❌ Não foi possível encontrar as colunas necessárias (NF, EMIT/FORN, MAT) na planilha.")
            df_base = None

    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo Excel: {e}")
        df_base = None
else:
    st.info("💡 Aguardando o upload da planilha base para ativar a validação de códigos.")

st.write("---")

# --- 2. SEÇÃO DE UPLOAD DOS XMLS ---
st.header("2️⃣ Processar Arquivos XML")
arquivos_xml = st.file_uploader("Escolha os arquivos XML da nota", type=["xml"], accept_multiple_files=True)

if arquivos_xml:
    if df_base is None:
        st.error("❌ Erro impeditivo: Você precisa carregar uma Planilha Base válida no Passo 1 antes de processar os XMLs!")
    else:
        cols_b = st.session_state.get('cols_base')
        
        for arquivo in arquivos_xml:
            try:
                conteudo_xml = arquivo.read()
                raiz = ET.fromstring(conteudo_xml)
                
                ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                infNFe = raiz.find('.//ns:infNFe', ns)
                if infNFe is None:
                    infNFe = raiz.find('.//infNFe')
                    ns = {}
                
                if infNFe is None:
                    st.error(f"O arquivo {arquivo.name} não é uma NF-e válida.")
                    continue
                    
                ide = infNFe.find('ns:ide', ns) if ns else infNFe.find('ide')
                num_nota = ide.find('ns:nNF', ns).text if ns else ide.find('nNF').text
                num_nota_limpo = str(int(num_nota)) if num_nota.isdigit() else str(num_nota).strip()
                
                # --- COLETAR FORNECEDOR (XML) ---
                fornecedor_final = "Não Identificado"
                emit = infNFe.find('ns:emit', ns) if ns else infNFe.find('emit')
                if emit is not None:
                    xNome = emit.find('ns:xNome', ns) if ns else emit.find('xNome')
                    xFant = emit.find('ns:xFant', ns) if ns else emit.find('xFant')
                    if xNome is not None: fornecedor_final = xNome.text
                    elif xFant is not None: fornecedor_final = xFant.text

                # --- COLETAR PESOS E VOLUMES (XML) ---
                transp = infNFe.find('ns:transp', ns) if ns else infNFe.find('transp')
                peso_liquido, peso_bruto, especie_volume, qtde_volume = 0.0, 0.0, "-", 0
                if transp is not None:
                    vol = transp.find('ns:vol', ns) if ns else transp.find('vol')
                    if vol is not None:
                        p_liq = vol.find('ns:pesoL', ns) if ns else vol.find('pesoL')
                        p_bru = vol.find('ns:pesoB', ns) if ns else vol.find('pesoB')
                        esp = vol.find('ns:esp', ns) if ns else vol.find('esp')
                        q_vol = vol.find('ns:qVol', ns) if ns else vol.find('qVol')
                        if p_liq is not None: peso_liquido = float(p_liq.text)
                        if p_bru is not None: peso_bruto =
