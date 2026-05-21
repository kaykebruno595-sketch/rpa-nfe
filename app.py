import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io

st.set_page_config(page_title="RPA - Extrator de XML", layout="centered")

st.title("📦 Extrator de Notas Fiscais (Acumulador Sem Erros)")
st.write("Crie e alimente a sua base de dados de XMLs de forma simples e direta.")

# 1. CAMPO PARA ALIMENTAR O HISTÓRICO (OPCIONAL)
st.subheader("1️⃣ Tem uma planilha dos dias anteriores? (Opcional)")
planilha_antiga = st.file_uploader("Se tiver, arraste a sua planilha antiga aqui para somar com os novos XMLs", type=["xlsx"])

df_historico = pd.DataFrame()
if planilha_antiga:
    try:
        df_historico = pd.read_excel(planilha_antiga)
        st.success(f"Planilha antiga carregada com sucesso! ({len(df_historico)} linhas encontradas).")
    except Exception as e:
        st.error(f"Erro ao ler a planilha antiga: {e}")

# 2. CAMPO PARA OS NOVOS XMLS
st.subheader("2️⃣ Arraste os NOVOS arquivos XML")
arquivos_xml = st.file_uploader("Escolha os novos arquivos XML do dia", type=["xml"], accept_multiple_files=True)

if arquivos_xml:
    dados_novos = []
    notas_ignoradas = 0
    
    # Listar notas que já existem na planilha antiga para evitar duplicados
    notas_existentes = []
    if not df_historico.empty and "Número da Nota" in df_historico.columns:
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
            
            # Trava antiduplicidade
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
        df_novos = pd.DataFrame(dados_novos)
        
        # Junta o histórico antigo com as novas notas
        if df_historico.empty:
            df_consolidado = df_novos
        else:
            df_consolidado = pd.concat([df_historico, df_novos], ignore_index=True)
            
        st.success(f"Sucesso! {len(df_novos)} novos itens processados.")
        if notas_ignoradas > 0:
            st.warning(f"{notas_ignoradas} nota(s) foram ignoradas por já estarem na sua planilha antiga.")
            
        # Mostra o resultado final na tela
        st.subheader("📋 Visualização da Nova Planilha Mãe")
        st.dataframe(df_consolidado)
        
        # Gerar o arquivo para download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_consolidado.to_excel(writer, index=False)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Baixar Planilha Mãe Atualizada (Excel)",
            data=buffer,
            file_name="CONSOLIDADO_NOTAS_GERAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
