import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import requests
import base64
import json

st.set_page_config(page_title="RPA - Extrator Acumulado", layout="centered")

st.title("📦 Extrator de Notas Fiscais (Histórico Único)")
st.write("Os novos XMLs arrastados serão somados à base de dados existente.")

# Configurações do seu GitHub
GITHUB_REPO = "kaykebruno59-sketch/rpa-nfe"
FILE_PATH = "planilha_acumulada.xlsx"
BRANCH = "principal"

# Recuperar o Token salvo no Secrets
if "GITHUB_TOKEN" not in st.secrets:
    st.error("Erro: O GITHUB_TOKEN não foi configurado no menu Secrets do Streamlit.")
    st.stop()

TOKEN = st.secrets["GITHUB_TOKEN"]
headers = {"Authorization": f"token {TOKEN}"}
url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{FILE_PATH}?ref={BRANCH}"

# Colunas padrão do nosso banco de dados
COLUNAS = ["Número da Nota", "Produto", "Quantidade", "Valor", "ICMS", "IPI", "Peso Líquido", "Peso Bruto"]

# --- FUNÇÃO PARA BUSCAR A PLANILHA EXISTENTE NO GITHUB ---
def buscar_planilha_github():
    try:
        resposta = requests.get(url, headers=headers)
        if resposta.status_code == 200:
            conteudo_json = resposta.json()
            conteudo_base64 = conteudo_json["content"]
            sha = conteudo_json["sha"]
            dados_binarios = base64.b64decode(conteudo_base64)
            # Tenta ler o Excel, se falhar (por estar em branco), cria um novo
            df = pd.read_excel(io.BytesIO(dados_binarios))
            return df, sha
        else:
            return pd.DataFrame(columns=COLUNAS), None
    except Exception:
        # Se der qualquer erro de leitura (arquivo corrompido ou em branco), recomeça do zero de forma segura
        return pd.DataFrame(columns=COLUNAS), None

# --- FUNÇÃO PARA SALVAR A PLANILHA ATUALIZADA NO GITHUB ---
def salvar_planilha_github(df_novo, sha_antigo):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_novo.to_excel(writer, index=False)
    buffer.seek(0)
    conteudo_base64 = base64.b64encode(buffer.read()).decode("utf-8")
    
    dados_envio = {
        "message": "🤖 RPA: Atualizando planilha de notas fiscais",
        "content": conteudo_base64,
        "branch": BRANCH
    }
    if sha_antigo:
        dados_envio["sha"] = sha_antigo
        
    resposta = requests.put(url, headers=headers, data=json.dumps(dados_envio))
    return resposta.status_code in [200, 201]

# Carregar o histórico atual de forma segura
df_historico, sha_atual = buscar_planilha_github()

# Mostrar o histórico atual na tela (se houver dados válidos)
if not df_historico.empty and len(df_historico.columns) == len(COLUNAS):
    st.subheader("📋 Histórico de Notas Já Salvas")
    st.dataframe(df_historico)

# Campo para o usuário arrastar os novos XMLs
arquivos_xml = st.file_uploader("Arraste os NOVOS arquivos XML aqui", type=["xml"], accept_multiple_files=True)

if arquivos_xml:
    dados_novos = []
    notas_ignoradas = 0
    
    # Pegar lista de notas que já existem para não duplicar (evita erro se o df estiver vazio)
    notas_existentes = []
    if "Número da Nota" in df_historico.columns:
        notas_existentes = df_historico["Número da Nota"].astype(str).tolist()
    
    for arquivo in arquivos_xml:
        try:
            conteudo_xml = arquivo.read()
            raiz = ET.fromstring(conteudo_xml)
            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
            
            infNFe = raiz.find('.//ns:infNFe', ns)
            if infNFe is None: continue
                
            ide = infNFe.find('ns:ide', ns)
            num_nota = ide.find('ns:nNF', ns).text if ide is not None else "N/A"
            
            # Evitar duplicados
            if str(num_nota) in notas_existentes:
                notas_ignoradas += 1
                continue
            
            # Pesos
            transp = infNFe.find('ns:transp', ns)
            peso_liquido = "-"
            peso_bruto = "-"
            if transp is not None:
                vol = transp.find('ns:vol', ns)
                if vol is not None:
                    p_liq = vol.find('ns:pesoL', ns)
                    p_bru = vol.find('ns:pesoB', ns)
                    if p_liq is not None: peso_liquido = p_liq.text
                    if p_bru is not None: peso_bruto = p_bru.text
            
            # Itens
            itens = infNFe.findall('ns:det', ns)
            for item in itens:
                prod = item.find('ns:prod', ns)
                nome_produto = prod.find('ns:xProd', ns).text
                quantidade = float(prod.find('ns:qCom', ns).text)
                valor_prod = float(prod.find('ns:vProd', ns).text)
                
                imposto = item.find('ns:imposto', ns)
                valor_icms = 0.0
                valor_ipi = 0.0
                if imposto is not None:
                    icms = imposto.find('.//ns:vICMS', ns)
                    ipi = imposto.find('.//ns:vIPI', ns)
                    if icms is not None: valor_icms = float(icms.text)
                    if ipi is not None: valor_ipi = float(ipi.text)
                
                dados_novos.append({
                    "Número da Nota": num_nota,
                    "Produto": nome_produto,
                    "Quantidade": quantidade,
                    "Valor": valor_prod,
                    "ICMS": valor_icms,
                    "IPI": valor_ipi,
                    "Peso Líquido": peso_liquido,
                    "Peso Bruto": peso_bruto
                })
        except Exception as e:
            st.error(f"Erro no arquivo {arquivo.name}: {e}")

    if dados_novos:
        df_novos_arquivos = pd.DataFrame(dados_novos)
        
        # Garante que o histórico tem as colunas certas antes de juntar
        if df_historico.empty or len(df_historico.columns) != len(COLUNAS):
            df_consolidado = df_novos_arquivos
        else:
            df_consolidado = pd.concat([df_historico, df_novos_arquivos], ignore_index=True)
        
        with st.spinner("Salvando e atualizando banco de dados..."):
            sucesso = salvar_planilha_github(df_consolidado, sha_atual)
            if sucesso:
                st.success(f"Sucesso! {len(df_novos_arquivos)} itens adicionados à planilha mãe.")
                if notas_ignoradas > 0:
                    st.warning(f"{notas_ignoradas} nota(s) foram ignoradas por já existirem no histórico.")
                st.rerun()
            else:
                st.error("Erro ao salvar os dados no GitHub. Verifique as permissões do seu Token.")

# Botão para baixar o Excel acumulado completo (se houver dados)
if not df_historico.empty and len(df_historico.columns) == len(COLUNAS):
    buffer_baixar = io.BytesIO()
    with pd.ExcelWriter(buffer_baixar, engine='openpyxl') as writer:
        df_historico.to_excel(writer, index=False)
    buffer_baixar.seek(0)
    
    st.download_button(
        label="📥 Baixar Planilha Mãe Completa (Excel)",
        data=buffer_baixar,
        file_name="CONSOLIDADO_NOTAS_GERAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
