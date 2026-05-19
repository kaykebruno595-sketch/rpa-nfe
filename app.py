import streamlit as st
import xml.etree.ElementTree as ET
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# --- CONEXÃO COM O GOOGLE SHEETS ---
def conectar_google_sheets():
    # Define o escopo de acesso às APIs do Google
    escopo = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    # Carrega a credencial do arquivo JSON
    if not os.path.exists("credenciais.json"):
        st.error("❌ Arquivo 'credenciais.json' não encontrado na pasta do servidor!")
        return None
        
    creds = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", escopo)
    cliente = gspread.authorize(creds)
    
    # ABRA PELO NOME EXATO DA SUA PLANILHA NO GOOGLE DRIVE
    # O e-mail do robô precisa estar como editor nela!
    planilha = cliente.open("Consolidado_Notas_Oficial").sheet1
    return planilha

# --- PROCESSAR O XML E SALVAR NA NUVEM ---
def processar_xml_web(arquivo_xml, planilha_google):
    # Ler o conteúdo do arquivo enviado pelo site
    conteudo_xml = arquivo_xml.read()
    root = ET.fromstring(conteudo_xml)

    ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
    prefixo = 'ns:' if ns else ''

    # 1. Pegar número da nota
    num_nota = ""
    elem_nNF = root.find(f".//{prefixo}ide/{prefixo}nNF", ns)
    if elem_nNF is not None:
        num_nota = elem_nNF.text

    # --- TRAVA ANTIDUPLICIDADE EM TEMPO REAL ---
    # Busca todos os valores da coluna A (Número da Nota) direto no Google Sheets
    try:
        notas_existentes = planilha_google.col_values(1) # Coluna 1 = Coluna A
        if str(num_nota) in notas_existentes:
            return f"⚠️ Nota {num_nota} já foi cadastrada anteriormente (Ignorada).", "aviso"
    except Exception as e:
        return f"Erro ao verificar duplicados: {e}", "erro"

    # 2. Se não for duplicada, extrai os pesos
    peso_liquido = 0.0
    peso_bruto = 0.0
    elem_peso_l = root.find(f".//{prefixo}transp/{prefixo}vol/{prefixo}pesoL", ns)
    if elem_peso_l is not None and elem_peso_l.text:
        peso_liquido = float(elem_peso_l.text)
        
    elem_peso_b = root.find(f".//{prefixo}transp/{prefixo}vol/{prefixo}pesoB", ns)
    if elem_peso_b is not None and elem_peso_b.text:
        peso_bruto = float(elem_peso_b.text)

    # 3. Extrai itens e envia linha por linha para a Nuvem
    itens = root.findall(f".//{prefixo}det", ns)
    linhas_para_inserir = []
    
    for item in itens:
        prod = item.find(f"{prefixo}prod", ns)
        nome_produto = prod.find(f"{prefixo}xProd", ns).text
        quantidade = float(prod.find(f"{prefixo}qCom", ns).text)
        valor_total_prod = float(prod.find(f"{prefixo}vProd", ns).text)
        
        imposto = item.find(f"{prefixo}imposto", ns)
        valor_icms = 0.0
        elem_icms = imposto.find(f".//{prefixo}vICMS", ns)
        if elem_icms is not None and elem_icms.text:
            valor_icms = float(elem_icms.text)
            
        valor_ipi = 0.0
        elem_ipi = imposto.find(f".//{prefixo}vIPI", ns)
        if elem_ipi is not None and elem_ipi.text:
            valor_ipi = float(elem_ipi.text)
            
        # Alinhado com o seu modelo do Google Sheets
        nova_linha = [
            num_nota, nome_produto, quantidade, valor_total_prod, 
            valor_icms, valor_ipi, peso_liquido, peso_bruto
        ]
        linhas_para_inserir.append(nova_linha)

    # Envia o bloco de linhas para o final da planilha do Google
    planilha_google.append_rows(linhas_para_inserir)
    return f"✅ Nota {num_nota} processada e enviada para o Google Sheets!", "sucesso"


# --- CONFIGURAÇÃO DA INTERFACE WEB (STREAMLIT) ---
st.set_page_config(page_title="RPA Notas Fiscais", page_icon="🤖")

st.title("🤖 Portal RPA - Consolidador de NF-e")
st.write("Insira os arquivos XML das Notas Fiscais abaixo. O sistema irá validar os dados e alimentar a planilha oficial na nuvem automaticamente.")

# Botão de Upload na Tela Web
arquivos = st.file_uploader("Arraste ou selecione os arquivos XML aqui", type="xml", accept_multiple_files=True)

if st.button("🚀 Processar Notas e Atualizar Nuvem"):
    if arquivos:
        # Conecta ao Google Sheets antes de começar o loop
        with st.spinner("Conectando ao Google Sheets..."):
            aba_google = conectar_google_sheets()
            
        if aba_google:
            # Passa por cada arquivo arrastado no site
            for arquivo in arquivos:
                mensagem, tipo = processar_xml_web(arquivo, aba_google)
                
                # Exibe o resultado individual de cada nota na tela do usuário
                if tipo == "sucesso":
                    st.success(mensagem)
                elif tipo == "aviso":
                    st.warning(mensagem)
                else:
                    st.error(mensagem)
            st.balloons() # Efeito visual de sucesso no site
    else:
        st.info("Por favor, selecione pelo menos um arquivo XML antes de clicar.")
