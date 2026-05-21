import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io

st.set_page_config(page_title="RPA - Extrator Instantâneo", layout="centered")

st.title("📦 RPA- Notas Fiscais")
st.write("Arraste os XMLs para extrair os dados na hora e alimentar seu modelo oficial.")

# Campo para novos XMLs
arquivos_xml = st.file_uploader("Arraste seus arquivos XML aqui", type=["xml"], accept_multiple_files=True)

if arquivos_xml:
    dados_novos = []
    
    for arquivo in arquivos_xml:
        try:
            conteudo_xml = arquivo.read()
            raiz = ET.fromstring(conteudo_xml)
            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
            
            infNFe = raiz.find('.//ns:infNFe', ns)
            if infNFe is None: continue
                
            ide = infNFe.find('ns:ide', ns)
            num_nota = ide.find('ns:nNF', ns).text if ide is not None else "N/A"
            
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
                
                # MODELO DE COLUNAS ESPECÍFICO DA SUA EMPRESA
                dados_novos.append({
                    "Número da Nota": num_nota,       # Coluna A
                    "Produto": nome_produto,         # Coluna B
                    "Quantidade": quantidade,        # Coluna C
                    "Valor": valor_prod,             # Coluna D
                    "ICMS": valor_icms,              # Coluna E
                    "IPI": valor_ipi,                # Coluna F
                    "Peso Líquido": peso_liquido,    # Coluna G
                    "Peso Bruto": peso_bruto         # Coluna H
                })
        except Exception as e:
            st.error(f"Erro no arquivo {arquivo.name}: {e}")

    if dados_novos:
        df_novos = pd.DataFrame(dados_novos)
        
        st.success(f"Sucesso! {len(df_novos)} itens processados na hora.")
        
        # Mostra a prévia na tela exatamente no layout do modelo
        st.subheader("📋 Dados Prontos (Modelo Oficial)")
        st.dataframe(df_novos)
        
        # Gerar o arquivo Excel temporário apenas com os dados novos
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_novos.to_excel(writer, index=False)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Baixar Dados Novos para Copiar",
            data=buffer,
            file_name="novos_dados_xml.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
