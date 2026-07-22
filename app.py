import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import time
import os
import io
import re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
try:
    from openpyxl.drawing.image import Image as ExcelImage
except ImportError:
    ExcelImage = None

# --- 1. CONFIGURAÇÃO VISUAL E IDENTIDADE ASA ---
st.set_page_config(page_title="ASA | Radar de Infraestrutura", page_icon="⚖️", layout="wide")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding-top: 2rem; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: 500; background-color: #6a9094; color: white; height: 3em; border: none; transition: all 0.3s ease; }
    .stButton>button:hover { background-color: #55787c; color: white; }
    div[data-testid="stMetricValue"] { color: #436468; font-size: 2.2rem; font-weight: 700; }
    h1, h2, h3, h4, h5, h6 { color: #436468 !important; font-weight: 600; }
    hr { border-bottom-color: #6a9094 !important; opacity: 0.3; }
    </style>
""", unsafe_allow_html=True)

if 'licitacoes_salvas' not in st.session_state: st.session_state['licitacoes_salvas'] = pd.DataFrame()
if 'resultados_busca' not in st.session_state: st.session_state['resultados_busca'] = pd.DataFrame()
if 'busca_realizada' not in st.session_state: st.session_state['busca_realizada'] = False

MAPA_MODALIDADES = {
    "Leilão": 1, "Diálogo Competitivo": 2, "Concurso": 3, 
    "Concorrência": 4, "Pregão Eletrônico": 6
}
LISTA_UFS = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"]

# --- FUNÇÃO DE ESTILIZAÇÃO DO EXCEL (COM LOGO) ---
def aplicar_estilo_excel(writer, df, sheet_name):
    worksheet = writer.sheets[sheet_name]
    
    cor_asa = PatternFill(start_color="436468", end_color="436468", fill_type="solid")
    fonte_branca = Font(color="FFFFFF", bold=True)
    borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # 1. Pintar Cabeçalho e aumentar a altura da linha 1 para caber a logo
    worksheet.row_dimensions[1].height = 45
    for cell in worksheet[1]:
        cell.fill = cor_asa
        cell.font = fonte_branca
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina
        
    # 2. Arrumar Larguras e Formatos das Colunas
    for col in worksheet.columns:
        col_letter = col[0].column_letter
        col_name = col[0].value
        
        if col_name == "Objeto":
            worksheet.column_dimensions[col_letter].width = 65
        elif col_name in ["Link", "Órgão"]:
            worksheet.column_dimensions[col_letter].width = 45
        elif col_name == "Status/Fase":
            worksheet.column_dimensions[col_letter].width = 25
        elif col_name == "Valor Estimado":
            worksheet.column_dimensions[col_letter].width = 22
        else:
            worksheet.column_dimensions[col_letter].width = 18
            
        for cell in col[1:]:
            cell.border = borda_fina
            if col_name == "Objeto":
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.alignment = Alignment(vertical="top")
                
            if col_name == "Valor Estimado" and isinstance(cell.value, (int, float)):
                cell.number_format = 'R$ #,##0.00'

    # 3. Adicionar a Logo no cantinho (Coluna J - Logo após o link)
    if ExcelImage is not None and os.path.exists("asa_logobrasao_verde.png"):
        try:
            img = ExcelImage("asa_logobrasao_verde.png")
            img.height = 50 # Altura ajustada
            img.width = 50  # Largura ajustada
            
            # Pinta a célula J1 de verde para dar continuidade ao cabeçalho
            worksheet['J1'].fill = cor_asa
            worksheet['J1'].border = borda_fina
            worksheet.column_dimensions['J'].width = 10
            
            # Insere a imagem flutuando na célula J1
            worksheet.add_image(img, 'J1')
        except Exception as e:
            pass # Se a imagem falhar, a planilha gera normalmente sem travar

# --- 2. MOTOR DE BUSCA E FILTROS ---
@st.cache_data(ttl=300)
def buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas):
    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
    todos_resultados = []
    erros = []
    if not modalidades_selecionadas: modalidades_selecionadas = ["Concorrência", "Leilão", "Diálogo Competitivo"]
        
    chunks = []
    atual = data_inicio
    while atual <= data_fim:
        proximo = min(atual + timedelta(days=15), data_fim)
        chunks.append((atual, proximo))
        atual = proximo + timedelta(days=1)
        
    for modalidade in modalidades_selecionadas:
        codigo = MAPA_MODALIDADES.get(modalidade)
        if not codigo: continue
        for inicio_chunk, fim_chunk in chunks:
            str_inicio = inicio_chunk.strftime("%Y%m%d")
            str_fim = fim_chunk.strftime("%Y%m%d")
            url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&codigoModalidadeContratacao={codigo}&pagina=1&tamanhoPagina=50"
            
            sucesso = False; tentativas = 0
            while not sucesso and tentativas < 3:
                try:
                    response = requests.get(url, headers=headers, timeout=30)
                    if response.status_code == 200:
                        todos_resultados.extend(response.json().get("data", []))
                        sucesso = True
                    else: tentativas += 1; time.sleep(2) 
                except: tentativas += 1; time.sleep(2)
            if not sucesso: erros.append(f"{modalidade} (bloco {str_inicio})")
            time.sleep(1.5)
    return todos_resultados, list(set(erros))

def filtrar_dados(licitacoes, palavras_chave, valor_min, valor_max, estados_selecionados):
    if not licitacoes: return pd.DataFrame()
    resultados = []
    lista_palavras = [p.strip().lower() for p in palavras_chave.split(',')] if palavras_chave else []
    ids_adicionados = set()
    
    for lic in licitacoes:
        id_unico = lic.get("id") or lic.get("linkSistemaOrigem")
        if id_unico in ids_adicionados: continue
        uf_licitacao = lic.get("unidadeOrgao", {}).get("ufSigla", "N/A")
        if estados_selecionados and (uf_licitacao not in estados_selecionados): continue
            
        objeto = lic.get("objetoCompra") or lic.get("sinteseObjeto") or lic.get("objeto") or "Descrição indisponível"
        objeto_str = str(objeto).lower()
        valor = lic.get("valorTotalEstimado") or 0.0 
        
        passou_palavra = any(p in objeto_str for p in lista_palavras) if (lista_palavras and lista_palavras[0] != "") else True
        if passou_palavra and (valor_min <= valor <= valor_max):
            numero_compra = lic.get("numeroCompra", "")
            ano_compra = lic.get("anoCompra", "")
            sequencial_compra = lic.get("sequencialCompra", "")
            cnpj = lic.get("orgaoEntidade", {}).get("cnpj", "")
            
            if cnpj and ano_compra and str(sequencial_compra) != "": 
                link_final = f"https://pncp.gov.br/app/editais/{cnpj}/{ano_compra}/{sequencial_compra}"
            elif lic.get("linkSistemaOrigem"):
                link_original = str(lic.get("linkSistemaOrigem", ""))
                link_final = link_original if link_original.startswith("http") else "https://" + link_original
            else: link_final = "https://pncp.gov.br"

            identificacao = f"{numero_compra}/{ano_compra}" if numero_compra and ano_compra else str(id_unico)
            fase = lic.get("situacaoCompraNome", "Não informada")

            resultados.append({
                "Identificação": identificacao,
                "Status/Fase": fase,
                "UF": uf_licitacao,
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "N/A"),
                "Modalidade": lic.get("modalidadeNome", "N/A"),
                "Objeto": objeto,
                "Valor Estimado": valor,
                "Data Publicação": lic.get("dataPublicacaoPncp"),
                "Link": link_final
            })
            ids_adicionados.add(id_unico)
    return pd.DataFrame(resultados)

# --- 3. FRONTEND ---
col_logo, col_titulo = st.columns([1, 8])
with col_logo:
    if os.path.exists("asa_logobrasao_verde.png"): st.image("asa_logobrasao_verde.png")
    else: st.markdown("<h2 style='text-align: center; color: #436468; margin-top: 15px;'>ASA</h2>", unsafe_allow_html=True)
with col_titulo:
    st.title("ASA - Radar Estratégico de Licitações")
    st.markdown("Monitoramento inteligente via **Portal Nacional de Contratações Públicas (PNCP)**.")

st.divider()

aba_busca, aba_interesse, aba_rastreador = st.tabs(["🔍 Nova Busca", "⭐ Unificador de Interesses", "📈 Rastreador via Planilha"])

# ==========================================
# ABA 1: BUSCA 
# ==========================================
with aba_busca:
    col_filtros, col_resultados = st.columns([1, 3], gap="large")
    with col_filtros:
        st.subheader("🎯 Parâmetros Regionais")
        estados_selecionados = st.multiselect("Estados de Interesse (UF):", LISTA_UFS)
        st.subheader("📅 Período")
        hoje = datetime.now()
        data_inicio = st.date_input("Data Inicial:", value=hoje - timedelta(days=15))
        data_fim = st.date_input("Data Final:", value=hoje)
        if data_inicio > data_fim: data_inicio, data_fim = data_fim, data_inicio
        st.subheader("🔎 Filtros Avançados")
        palavras_chave = st.text_input("Palavras-chave (separadas por vírgula):", "")
        modalidades_selecionadas = st.multiselect("Modalidades:", list(MAPA_MODALIDADES.keys()), default=["Concorrência", "Leilão"])
        st.subheader("💰 Filtro Financeiro")
        valor_min = st.number_input("Valor Mín. (R$):", value=100000000.0, step=1000000.0)
        valor_max = st.number_input("Valor Máx. (R$):", value=5000000000.0, step=1000000.0)
        buscar = st.button("🔍 Mapear Oportunidades")

    with col_resultados:
        if buscar:
            with st.spinner("Varrendo os servidores do Governo..."):
                dados_brutos, erros = buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas)
                if erros: st.warning(f"Avisos de rede: {', '.join(erros)}")
                if len(dados_brutos) > 0:
                    df_final = filtrar_dados(dados_brutos, palavras_chave, valor_min, valor_max, estados_selecionados)
                    st.session_state['resultados_busca'] = df_final
                    st.session_state['busca_realizada'] = True
                else:
                    st.session_state['resultados_busca'] = pd.DataFrame()
                    st.session_state['busca_realizada'] = True
                    st.info("Nenhuma licitação encontrada ou bloqueio temporário do governo.")

        if st.session_state['busca_realizada']:
            df_atual = st.session_state['resultados_busca']
            if not df_atual.empty:
                if "Acompanhar" not in df_atual.columns: df_atual.insert(0, "Acompanhar", False)
                st.write("### 📌 Resultados da Busca")
                m1, m2 = st.columns(2)
                m1.metric("Encontradas (com filtro)", f"{len(df_atual)}")
                m2.metric("Volume Financeiro", f"R$ {df_atual['Valor Estimado'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                if "Data Publicação" in df_atual.columns: df_atual = df_atual.sort_values(by="Data Publicação", ascending=False)
                
                df_editado = st.data_editor(
                    df_atual,
                    column_config={"Acompanhar": st.column_config.CheckboxColumn("⭐ Salvar", default=False), "Link": st.column_config.LinkColumn("Edital", display_text="Acessar 🔗")},
                    disabled=["Identificação", "Status/Fase", "UF", "Órgão", "Modalidade", "Objeto", "Valor Estimado", "Data Publicação"],
                    hide_index=True, use_container_width=True, key="editor_resultados"
                )
                
                if st.button("💾 Enviar selecionadas para Aba de Interesse"):
                    selecionadas = df_editado[df_editado["Acompanhar"] == True].copy().drop(columns=["Acompanhar"])
                    if not selecionadas.empty:
                        st.session_state['licitacoes_salvas'] = pd.concat([st.session_state['licitacoes_salvas'], selecionadas]).drop_duplicates(subset=["Identificação"])
                        st.success("Licitações salvas com sucesso! Vá para a aba de Interesse.")
                    else: st.warning("Você não marcou nenhuma licitação.")
            else: st.warning("Nenhuma licitação passou nos filtros.")

# ==========================================
# ABA 2: UNIFICADOR DE INTERESSES (Upload da Base)
# ==========================================
with aba_interesse:
    st.subheader("⭐ Seu Painel de Acompanhamento (Unificador)")
    st.markdown("Suba sua Planilha Mestre (opcional) para juntar com as novas pesquisas de hoje e baixe o arquivo atualizado.")
    
    arquivo_base = st.file_uploader("📂 Upload da Planilha Mestre (.xlsx)", type=["xlsx"], key="up_mestre")
    df_export = st.session_state['licitacoes_salvas'].copy()

    if arquivo_base is not None:
        try:
            df_antigo = pd.read_excel(arquivo_base)
            df_export = pd.concat([df_antigo, df_export]).drop_duplicates(subset=["Identificação"], keep="last")
            st.success("✅ Planilha antiga carregada e unida com sucesso!")
        except Exception as e:
            st.error(f"Erro ao ler a planilha: {e}")

    if df_export.empty:
        st.info("Nenhuma licitação salva no momento.")
    else:
        st.write("**Pré-visualização do Arquivo Final:**")
        st.dataframe(df_export.style.format({"Valor Estimado": "R$ {:,.2f}"}), hide_index=True, use_container_width=True)
        
        buffer_exp = io.BytesIO()
        with pd.ExcelWriter(buffer_exp, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Base ASA')
            aplicar_estilo_excel(writer, df_export, 'Base ASA')
            
        st.download_button(
            label="📥 Baixar Planilha Mestre Atualizada (.xlsx)",
            data=buffer_exp.getvalue(),
            file_name=f"Master_ASA_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.divider()
        if st.button("🗑️ Limpar Lista Temporária"):
            st.session_state['licitacoes_salvas'] = pd.DataFrame()
            st.rerun()

# ==========================================
# ABA 3: RASTREADOR VIA UPLOAD
# ==========================================
with aba_rastreador:
    st.subheader("📈 Rastreador de Status via Planilha")
    st.markdown("Suba sua Planilha Mestre aqui. O robô vai ler a coluna **Link**, consultar o governo para descobrir a fase atual de cada processo e te devolver um Excel novinho.")
    
    arquivo_rastreio = st.file_uploader("📂 Upload da Planilha para Rastreio (.xlsx)", type=["xlsx"], key="up_rastreio")
    
    if arquivo_rastreio is not None:
        try:
            df_rastrear = pd.read_excel(arquivo_rastreio)
            if "Link" not in df_rastrear.columns:
                st.error("❌ A planilha enviada não possui uma coluna chamada 'Link'. Verifique o arquivo e tente novamente.")
            else:
                if st.button("🔄 Rastrear e Atualizar Todos os Status"):
                    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
                    progresso = st.progress(0)
                    total = len(df_rastrear)
                    
                    with st.spinner("Lendo banco de dados do Governo..."):
                        for index, row in df_rastrear.iterrows():
                            link = str(row["Link"]).strip()
                            try:
                                match = re.search(r"editais/(\d+)/(\d+)/(\d+)", link)
                                if match:
                                    cnpj, ano, seq = match.groups()
                                    url_status = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}"
                                    resp = requests.get(url_status, headers=headers, timeout=15)
                                    
                                    if resp.status_code == 200:
                                        dados_lic = resp.json()
                                        df_rastrear.at[index, "Status/Fase"] = dados_lic.get("situacaoCompraNome", "Não informada")
                                time.sleep(0.5) 
                            except:
                                pass
                            
                            progresso.progress((index + 1) / total)
                    
                    st.success("✅ Varredura concluída! Confira abaixo os status atualizados.")
                    st.dataframe(df_rastrear.style.format({"Valor Estimado": "R$ {:,.2f}"}), hide_index=True, use_container_width=True)
                    
                    buffer_rastreio = io.BytesIO()
                    with pd.ExcelWriter(buffer_rastreio, engine='openpyxl') as writer:
                        df_rastrear.to_excel(writer, index=False, sheet_name='Base Atualizada ASA')
                        aplicar_estilo_excel(writer, df_rastrear, 'Base Atualizada ASA')
                    
                    st.download_button(
                        label="📥 Baixar Planilha com Status Atualizado (.xlsx)",
                        data=buffer_rastreio.getvalue(),
                        file_name=f"Master_ASA_Atualizada_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Ocorreu um erro na leitura do arquivo: {e}")
