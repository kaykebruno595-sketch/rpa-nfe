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
st.write("Atualize a planilha base e converta seus XMLs direto por aqui!")

# --- FUNÇÃO PARA LIMPEZA EXTREMA DE TEXTO ---
def limpar_texto_comparacao(texto):
    if pd.isna(texto):
        return ""
    txt = str(texto).upper().strip()
    txt = re.sub(r'[.\-\/]', '', txt)
    txt = txt.replace("LTDA", "").replace("S/A", "").replace("SA", "").replace("S.A", "")
    return "".join(txt.split())

# --- 1. SEÇÃO DE UPLOAD DA PLANILHA BASE ---
st.header("1️⃣ Atualizar Planilha Base (Trânsito)")
arquivo_base_upload = st.file_uploader("Arraste aqui a planilha 'base_transito.xlsx' atualizada", type=["xlsx"])

df_base = None

if arquivo_base_upload is not None:
    try:
        # Lê a planilha que o usuário acabou de arrastar no site
        df_base = pd.read_excel(arquivo_base_upload, sheet_name=0)
        
        # Padroniza os nomes das colunas para Maiúsculo
        df_base.columns = [str(col).strip().upper() for col in df_base.columns]
        
        st.success("✅ Planilha Base carregada e ativa!")
        
        # --- PAINEL DE CERTEZA VISUAL ---
        st.subheader("👀 Dados Ativos no Momento (Últimas 5 linhas):")
        col_nf = next((c for c in df_base.columns if "NF" in c), None)
        col_emit = next((c for c in df_base.columns if "EMIT" in c or "FORN" in c), None)
        col_mat = next((c for c in df_base.columns if "MAT" in c), None)
        
        if col_nf and col_emit and col_mat:
            df_preview = df_base[[col_nf, col_emit, col_mat]].tail(5)
            df_preview.columns = ["Nota Fiscal (nF)", "Fornecedor (emitNome)", "Código Certo (id_material)"]
            st.dataframe(df_preview, use_container_width=True)
        else:
            st.warning("⚠️ Atenção: Verifique se a sua planilha possui as colunas `nF`, `emitNome` e `id_material`.")
            st.write("Colunas detectadas:", list(df_base.columns))
            
    except Exception as e:
        st.error(f"Erro ao processar a planilha base: {e}")
else:
    st.info("💡 Aguardando o upload da planilha base para ativar a validação de códigos.")

st.write("---")

# --- 2. SEÇÃO DE UPLOAD DOS XMLS ---
st.header("2️⃣ Processar Arquivos XML")
arquivos_xml = st.file_uploader("Escolha os arquivos XML da nota", type=["xml"], accept_multiple_files=True)

if arquivos_xml:
    if df_base is None:
        st.error("❌ Erro: Você precisa carregar a Planilha Base no passo 1 antes de processar os XMLs!")
    else:
        for arquivo in arquivos_xml:
            try:
                conteudo_xml = arquivo.read()
                raiz = ET.fromstring(conteudo_xml)
                ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                
                infNFe = raiz.find('.//ns:infNFe', ns)
                if infNFe is None:
                    st.error(f"O arquivo {arquivo.name} não é uma NF-e válida.")
                    continue
                    
                ide = infNFe.find('ns:ide', ns)
                num_nota = ide.find('ns:nNF', ns).text if ide is not None else "N/A"
                num_nota_limpo = str(int(num_nota)) if num_nota.isdigit() else str(num_nota).strip()
                
                # --- COLETAR FORNECEDOR (XML) ---
                fornecedor_final = "Não Identificado"
                emit = infNFe.find('ns:emit', ns)
                if emit is not None:
                    xNome = emit.find('ns:xNome', ns)
                    xFant = emit.find('ns:xFant', ns)
                    if xNome is not None: fornecedor_final = xNome.text
                    elif xFant is not None: fornecedor_final = xFant.text

                # --- COLETAR CIDADE COM TRAVA CARIACICA (XML) ---
                cidade_final = "Outros / Não Encontrado"
                dest = infNFe.find('ns:dest', ns)
                if dest is not None:
                    enderDest = dest.find('ns:enderDest', ns)
                    if enderDest is not None:
                        xMun_node = enderDest.find('ns:xMun', ns)
                        if xMun_node is not None:
                            cidade_xml_bruta = xMun_node.text.upper()
                            if "CARIACICA" in cidade_xml_bruta: cidade_final = "Positive CO"
                            elif "NATAL" in cidade_xml_bruta: cidade_final = "Natal"
                            elif "POSITIVE" in cidade_xml_bruta: cidade_final = "Positive"
                            elif "SANTA LUZIA" in cidade_xml_bruta: cidade_final = "Santa Luzia"
                            elif "ARAMA" in cidade_xml_bruta: cidade_final = "Arama"
                            elif "LONDRINA" in cidade_xml_bruta: cidade_final = "Londrina"
                            elif "DIADEMA" in cidade_xml_bruta: cidade_final = "Diadema"
                            else: cidade_final = cidade_xml_bruta.title()

                # --- COLETAR PESOS E VOLUMES (XML) ---
                transp = infNFe.find('ns:transp', ns)
                peso_liquido, peso_bruto, especie_volume, qtde_volume = 0.0, 0.0, "-", 0
                if transp is not None:
                    vol = transp.find('ns:vol', ns)
                    if vol is not None:
                        p_liq = vol.find('ns:pesoL', ns)
                        p_bru = vol.find('ns:pesoB', ns)
                        esp = vol.find('ns:esp', ns)
                        q_vol = vol.find('ns:qVol', ns)
                        if p_liq is not None: peso_liquido = float(p_liq.text)
                        if p_bru is not None: peso_bruto = float(p_bru.text)
                        if esp is not None: especie_volume = esp.text
                        if q_vol is not None: qtde_volume = int(q_vol.text)

                # --- CRUZAMENTO DE DADOS COM A PLANILHA CARREGADA ---
                codigo_substituto = None
                col_nf_base = next((c for c in df_base.columns if "NF" in c), None)
                col_emit_base = next((c for c in df_base.columns if "EMIT" in c or "FORN" in c), None)
                col_mat_base = next((c for c in df_base.columns if "MAT" in c), None)
                
                if col_nf_base and col_emit_base and col_mat_base:
                    fornecedor_xml_ultra_limpo = limpar_texto_comparacao(fornecedor_final)
                    
                    for _, linha_base in df_base.iterrows():
                        nota_base_str = str(linha_base[col_nf_base]).split('.')[0].strip()
                        nota_base_limpa = str(int(nota_base_str)) if nota_base_str.isdigit() else nota_base_str
                        
                        if nota_base_limpa == num_nota_limpo:
                            fornecedor_base_ultra_limpo = limpar_texto_comparacao(linha_base[col_emit_base])
                            if fornecedor_base_ultra_limpo in fornecedor_xml_ultra_limpo or fornecedor_xml_ultra_limpo in fornecedor_base_ultra_limpo or fornecedor_base_ultra_limpo[:8] in fornecedor_xml_ultra_limpo:
                                codigo_substituto = str(linha_base[col_mat_base]).strip()
                                break

                # --- MONTAGEM DOS ITENS ---
                lista_produtos = []
                itens_xml = infNFe.findall('ns:det', ns)
                
                for item in itens_xml:
                    prod = item.find('ns:prod', ns)
                    if codigo_substituto and codigo_substituto.lower() != 'nan':
                        codigo_final_item = codigo_substituto
                    else:
                        codigo_final_item = prod.find('ns:cProd', ns).text
                    
                    nome_produto = prod.find('ns:xProd', ns).text
                    umb = prod.find('ns:uCom', ns).text  
                    quantidade = float(prod.find('ns:qCom', ns).text)
                    valor_unitario = float(prod.find('ns:vUnCom', ns).text)
                    valor_total_item = float(prod.find('ns:vProd', ns).text)
                    
                    # Captura os impostos como números reais/decimais para formatação posterior
                    imposto = item.find('ns:imposto', ns)
                    valor_icms_num, valor_ipi_num = 0.0, 0.0
                    if imposto is not None:
                        icms_detalhe = imposto.find('.//ns:pICMS', ns)
                        if icms_detalhe is not None: valor_icms_num = float(icms_detalhe.text)
                        ipi_detalhe = imposto.find('.//ns:pIPI', ns)
                        if ipi_detalhe is not None: valor_ipi_num = float(ipi_detalhe.text)
                    
                    lista_produtos.append({
                        "FORNECEDOR": fornecedor_final,
                        "CIDADE/MUNICÍPIO": cidade_final,
                        "CODIGO": codigo_final_item,
                        "DESCRIÇÃO": nome_produto,
                        "NOTA FISCAL": num_nota,
                        "UMB": umb,
                        "QTDE": quantidade,
                        "VLR. UNT.": valor_unitario,
                        "VLR. TT.": valor_total_item,
                        "ICMS": valor_icms_num,
                        "IPI": valor_ipi_num
                    })

                # --- FORMATAÇÃO DO ARQUIVO EXCEL FINAL ---
                wb = Workbook()
                ws = wb.active
                ws.title = f"NF {num_nota}"
                ws.views.sheetView[0].showGridLines = True
                
                cor_azul_escuro, cor_azul_claro = "1B365D", "F0F4F8"
                fill_header = PatternFill(start_color=cor_azul_escuro, end_color=cor_azul_escuro, fill_type="solid")
                fill_sub_header = PatternFill(start_color=cor_azul_claro, end_color=cor_azul_claro, fill_type="solid")
                font_branca_negrito = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                font_preta_negrito = Font(name="Calibri", size=11, bold=True, color="000000")
                font_normal = Font(name="Calibri", size=11)
                
                border_fina = Border(
                    left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
                    top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
                )
                
                ws.merge_cells("A1:K1")
                ws["A1"] = "DADOS MATERIAIS"
                ws["A1"].fill = fill_header
                ws["A1"].font = font_branca_negrito
                ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
                ws.row_dimensions[1].height = 25
                
                colunas = ["FORNECEDOR", "CIDADE/MUNICÍPIO", "CÓDIGO", "DESCRIÇÃO", "NOTA FISCAL", "UMB", "QTDE", "VLR. UNT.", "VLR. TT.", "ICMS", "IPI"]
                for col_idx, texto_coluna in enumerate(colunas, 1):
                    celula = ws.cell(row=2, column=col_idx, value=texto_coluna)
                    celula.fill = fill_header
                    celula.font = font_branca_negrito
                    celula.alignment = Alignment(horizontal="center", vertical="center")
                    celula.border = border_fina
                ws.row_dimensions[2].height = 22
                
                linha_atual = 3
                for prod in lista_produtos:
                    ws.cell(row=linha_atual, column=1, value=prod["FORNECEDOR"]).alignment = Alignment(horizontal="left")
                    ws.cell(row=linha_atual, column=2, value=prod["CIDADE/MUNICÍPIO"]).alignment = Alignment(horizontal="left")
                    ws.cell(row=linha_atual, column=3, value=prod["CODIGO"]).alignment = Alignment(horizontal="center")
                    ws.cell(row=linha_atual, column=4, value=prod["DESCRIÇÃO"]).alignment = Alignment(horizontal="left")
                    ws.cell(row=linha_atual, column=5, value=int(prod["NOTA FISCAL"])).alignment = Alignment(horizontal="center")
                    ws.cell(row=linha_atual, column=6, value=prod["UMB"]).alignment = Alignment(horizontal="center")
                    
                    # Formatação numérica pura (sem siglas de moedas ou símbolos de porcentagem)
                    ws.cell(row=linha_atual, column=7, value=prod["QTDE"]).number_format = '#,##0.00'
                    
                    # Alteração solicitada: Removido 'R$' e configurado com 2 casas decimais
                    ws.cell(row=linha_atual, column=8, value=prod["VLR. UNT."]).number_format = '#,##0.00'
                    ws.cell(row=linha_atual, column=9, value=prod["VLR. TT."]).number_format = '#,##0.00'
                    
                    # Alteração solicitada: Removido '%' e configurado com 2 casas decimais
                    celula_icms = ws.cell(row=linha_atual, column=10, value=prod["ICMS"])
                    celula_icms.number_format = '#,##0.00'
                    celula_icms.alignment = Alignment(horizontal="center")
                    
                    celula_ipi = ws.cell(row=linha_atual, column=11, value=prod["IPI"])
                    celula_ipi.number_format = '#,##0.00'
                    celula_ipi.alignment = Alignment(horizontal="center")
                    
                    for c in range(1, 12):
                        ws.cell(row=linha_atual, column=c).font = font_normal
                        ws.cell(row=linha_atual, column=c).border = border_fina
                    ws.row_dimensions[linha_atual].height = 20
                    linha_atual += 1
                    
                linha_atual += 1
                ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=3)
                celula_fiscal_titulo = ws.cell(row=linha_atual, column=1, value="DADOS FISCAIS")
                celula_fiscal_titulo.fill = fill_header
                celula_fiscal_titulo.font = font_branca_negrito
                ws.row_dimensions[linha_atual].height = 22
                linha_atual += 1
                
                dados_fiscais_valores = [
                    ("PESO BRUTO", peso_bruto),
                    ("PESO LÍQUIDO", peso_liquido),
                    ("ESPÉCIE DE VOLUME", especie_volume),
                    ("QTDE VOLUME", qtde_volume)
                ]
                
                for label, valor in dados_fiscais_valores:
                    ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=2)
                    c_label = ws.cell(row=linha_atual, column=1, value=label)
                    c_label.fill = fill_sub_header
                    c_label.font = font_preta_negrito
                    c_label.border = border_fina
                    
                    c_valor = ws.cell(row=linha_atual, column=3, value=valor)
                    if isinstance(valor, float): c_valor.number_format = '#,##0.00'
                    c_valor.font = font_normal
                    c_valor.border = border_fina
                    c_valor.alignment = Alignment(horizontal="left")
                    
                    ws.cell(row=linha_atual, column=2).border = border_fina
                    ws.row_dimensions[linha_atual].height = 18
                    linha_atual += 1
                    
                for col in ws.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.row == 1: continue  
                        if cell.value: max_len = max(max_len, len(str(cell.value)))
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
                
                ws.column_dimensions['A'].width = 32
                ws.column_dimensions['B'].width = 22
                ws.column_dimensions['D'].width = 45
                
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                fornecedor_limpo = re.sub(r'[\\/*?:"<>|]', "", fornecedor_final).strip()
                status_origem = "🏷️ [ID Alterado com Sucesso]" if codigo_substituto else "⚠️ [Nota Não Encontrada na Base - Mantido XML]"
                
                st.download_button(
                    label=f"📥 Baixar nota {num_nota} - {fornecedor_limpo} {status_origem}",
                    data=buffer,
                    file_name=f"PLANILHA_NOTA_{num_nota}_{fornecedor_limpo}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Erro ao processar o arquivo {arquivo.name}: {e}")
