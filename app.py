import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="RPA - Gerador de Planilha de Nota", layout="centered")

st.title("📦 Extrator Xml Copacker")

# --- FUNÇÃO PARA LIMPEZA EXTREMA DE TEXTO (EVITAR ERROS DE MATCH) ---
def limpar_texto_comparacao(texto):
    if pd.isna(texto):
        return ""
    txt = str(texto).upper().strip()
    # Remove acentos e caracteres especiais comuns
    txt = re.sub(r'[.\-\/]', '', txt)
    # Remove sufixos empresariais comuns que mudam entre sistemas
    txt = txt.replace("LTDA", "").replace("S/A", "").replace("SA", "").replace("S.A", "")
    return "".join(txt.split()) # Remove todos os espaços em branco

# --- CARREGAMENTO AUTOMÁTICO DA PLANILHA BASE DO GITHUB ---
nome_arquivo_base = "base_transito.xlsx"
df_base = None

if os.path.exists(nome_arquivo_base):
    try:
        # Lê a planilha base (tenta pegar a aba 'Trânsito', se não achar pega a primeira)
        try:
            df_base = pd.read_excel(nome_arquivo_base, sheet_name="Trânsito")
        except:
            df_base = pd.read_excel(nome_arquivo_base, sheet_name=0)
            
        # Padroniza os nomes das colunas
        df_base.columns = [str(col).strip().upper() for col in df_base.columns]
        
        st.success("✅ Planilha Base carregada com sucesso do repositório GitHub!")
        
        # --- PAINEL DE CERTEZA VISUAL ---
        st.subheader("👀 Monitor da Planilha Base")
        st.write("Confira abaixo os últimos dados que o sistema acabou de ler (Últimas 5 linhas):")
        
        # Mapeia as colunas reais para exibição amigável
        col_nf = next((c for c in df_base.columns if "NF" in c), None)
        col_emit = next((c for c in df_base.columns if "EMIT" in c or "FORN" in c), None)
        col_mat = next((c for c in df_base.columns if "MAT" in c), None)
        
        if col_nf and col_emit and col_mat:
            df_preview = df_base[[col_nf, col_emit, col_mat]].tail(5)
            df_preview.columns = ["Nota Fiscal (nF)", "Fornecedor (emitNome)", "Código Certo (id_material)"]
            st.dataframe(df_preview, use_container_width=True)
        else:
            st.warning("⚠️ Suas colunas não estão com os nomes exatos: `nF`, `emitNome` e `id_material`.")
            st.write("Colunas encontradas na sua planilha:", list(df_base.columns))
            
    except Exception as e:
        st.error(f"Erro ao ler a planilha base '{nome_arquivo_base}': {e}")
else:
    st.warning(f"⚠️ Arquivo '{nome_arquivo_base}' não encontrado no repositório. O sistema usará os códigos nativos do XML.")

st.write("---")
# Campo para fazer upload dos XMLs do dia a dia
arquivos_xml = st.file_uploader("Escolha os arquivos XML da nota", type=["xml"], accept_multiple_files=True)

if arquivos_xml:
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
            
            # Limpeza do número da nota para comparação segura (ex: tira zeros à esquerda)
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

            # --- CORREÇÃO CIRÚRGICA: BUSCA INTELIGENTE NA TABELA BASE ---
            codigo_substituto = None
            
            if df_base is not None:
                # Localiza as colunas ideais independente da caixa alta/baixa
                col_nf_base = next((c for c in df_base.columns if "NF" in c), None)
                col_emit_base = next((c for c in df_base.columns if "EMIT" in c or "FORN" in c), None)
                col_mat_base = next((c for c in df_base.columns if "MAT" in c), None)
                
                if col_nf_base and col_emit_base and col_mat_base:
                    fornecedor_xml_ultra_limpo = limpar_texto_comparacao(fornecedor_final)
                    
                    for _, linha_base in df_base.iterrows():
                        # Limpa a nota da linha atual da planilha
                        nota_base_str = str(linha_base[col_nf_base]).split('.')[0].strip()
                        nota_base_limpa = str(int(nota_base_str)) if nota_base_str.isdigit() else nota_base_str
                        
                        # Se o número da nota bater...
                        if nota_base_limpa == num_nota_limpo:
                            # Limpa o fornecedor da linha atual da planilha para comparar
                            fornecedor_base_ultra_limpo = limpar_texto_comparacao(linha_base[col_emit_base])
                            
                            # Se um fornecedor contiver o outro após a limpeza pesada, é o match perfeito!
                            if fornecedor_base_ultra_limpo in fornecedor_xml_ultra_limpo or proveedores_xml_ultra_limpo in fornecedor_base_ultra_limpo or fornecedor_base_ultra_limpo[:8] in fornecedor_xml_ultra_limpo:
                                codigo_substituto = str(linha_base[col_mat_base]).strip()
                                break

            # --- COLETAR OS ITENS DO XML APLICANDO AS REGRAS ---
            lista_produtos = []
            itens_xml = infNFe.findall('ns:det', ns)
            
            for item in itens_xml:
                prod = item.find('ns:prod', ns)
                
                # Se a validação cruzada funcionou, aplica o id_material. Senão, mantém o do XML.
                if codigo_substituto and codigo_substituto.lower() != 'nan':
                    codigo_final_item = codigo_substituto
                else:
                    codigo_final_item = prod.find('ns:cProd', ns).text
                
                nome_produto = prod.find('ns:xProd', ns).text
                umb = prod.find('ns:uCom', ns).text  
                quantidade = float(prod.find('ns:qCom', ns).text)
                valor_unitario = float(prod.find('ns:vUnCom', ns).text)
                valor_total_item = float(prod.find('ns:vProd', ns).text)
                
                imposto = item.find('ns:imposto', ns)
                valor_icms_penc, valor_ipi_penc = "0%", "0%"
                if imposto is not None:
                    icms_detalhe = imposto.find('.//ns:pICMS', ns)
                    if icms_detalhe is not None: valor_icms_penc = f"{int(float(icms_detalhe.text))}%"
                    ipi_detalhe = imposto.find('.//ns:pIPI', ns)
                    if ipi_detalhe is not None: valor_ipi_penc = f"{int(float(ipi_detalhe.text))}%"
                
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
                    "ICMS": valor_icms_penc,
                    "IPI": valor_ipi_penc
                })

            # --- CONSTRUÇÃO DA PLANILHA NO EXCEL ---
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
                
                ws.cell(row=linha_atual, column=7, value=prod["QTDE"]).number_format = '#,##0.00'
                ws.cell(row=linha_atual, column=8, value=prod["VLR. UNT."]).number_format = 'R$ #,##0.00'
                ws.cell(row=linha_atual, column=9, value=prod["VLR. TT."]).number_format = 'R$ #,##0.00'
                
                ws.cell(row=linha_atual, column=10, value=str(prod["ICMS"])).alignment = Alignment(horizontal="center")
                ws.cell(row=linha_atual, column=11, value=str(prod["IPI"])).alignment = Alignment(horizontal="center")
                
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
            status_origem = "🏷️ [ID Verificado pela Base]" if codigo_substituto else "⚠️ [Nota/Fornecedor Não Encontrado na Base - Usando XML]"
            
            st.download_button(
                label=f"📥 Baixar nota {num_nota} - {fornecedor_limpo} {status_origem}",
                data=buffer,
                file_name=f"PLANILHA_NOTA_{num_nota}_{fornecedor_limpo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Erro ao processar o arquivo {arquivo.name}: {e}")
