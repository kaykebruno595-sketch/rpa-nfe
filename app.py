import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="RPA - Gerador de Planilha de Nota", layout="centered")

st.title("📦 RPA Conversor de XML Copacker")
st.write("Atualize a planilha base e converta seus XMLs!")

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
        # Limpa floats esquisitos vindos do Excel na coluna de NF (Ex: 123.0 -> 123)
        df_base[col_nf] = df_base[col_nf].apply(lambda x: str(int(float(x))) if re.match(r'^\d+\.\d+$', str(x)) else str(x).strip())
        
        st.success("✅ Planilha Base carregada e ativa na memória do site!")
        
        colunas_exibicao = [col_nf, col_emit, col_mat]
        if col_desc_base and col_desc_base not in colunas_exibicao:
            colunas_exibicao.append(col_desc_base)
            
        df_preview = df_base[colunas_exibicao].tail(5)
        st.dataframe(df_preview, use_container_width=True)
        
        # Corrigido: Alinhado corretamente dentro do 'if' e usando 'col_mat'
        st.session_state['cols_base'] = {'nf': col_nf, 'emit': col_emit, 'mat': col_mat}
