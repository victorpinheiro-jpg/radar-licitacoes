import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import time
import os
import io
import re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

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

# --- CÁLCULO DE DIAS RESTANTES ---
def calcular_dias_restantes(data_sessao_str):
    if not data_sessao_str or data_sessao_str == "Verificar Edital":
        return "N/A"
    try:
        dt_sessao = datetime.strptime(data_sessao_str, '%d/%m/%Y %H:%M')
        hoje = datetime.now()
        delta = (dt_sessao.date() - hoje.date()).days
        if delta < 0: return "Sessão Encerrada"
        elif delta == 0: return "🚨 É HOJE!"
        else: return f"Faltam {delta} dias"
    except:
        return "N/A"

# --- FUNÇÃO DE ESTILIZAÇÃO DO EXCEL (Layout Premium + Alertas) ---
def aplicar_estilo_excel(writer, df, sheet_name):
    worksheet = writer.sheets[sheet_name]
    
    # Paleta de Cores
    cor_asa = PatternFill(start_color="436468", end_color="436468", fill_type="solid")
    cor_fundo_claro = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
    cor_suspensa = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Amarelo Alerta
    cor_morta = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid") # Cinza Inativo
    
    fonte_branca = Font(color="FFFFFF", bold=True)
    fonte_normal = Font(name="Calibri", size=11)
    fonte_urgente = Font(name="Calibri", size=11, color="FF0000", bold=True) # Texto Vermelho
    
    borda_fina = Border(left=Side(style='thin', color="BFBFBF"), right=Side(style='thin', color="BFBFBF"), 
                        top=Side(style='thin', color="BFBFBF"), bottom=Side(style='thin', color="BFBFBF"))
    
    # Cabeçalho
    worksheet.row_dimensions[1].height = 30
    col_indices = {cell.value: idx + 1 for idx, cell in enumerate(worksheet[1])}
    status_col_idx = col_indices.get("Status/Fase")
    dias_col_idx = col_indices.get("Dias Restantes")

    for cell in worksheet[1]:
        cell.fill = cor_asa
        cell.font = fonte_branca
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina
        
    # Larguras das Colunas
    for col in worksheet.columns:
        col_letter = col[0].column_letter
        col_name = col[0].value
        
        if col_name == "Objeto": worksheet.column_dimensions[col_letter].width = 65
        elif col_name == "Órgão": worksheet.column_dimensions[col_letter].width = 40
        elif col_name == "Link": worksheet.column_dimensions[col_letter].width = 35
        elif col_name == "Anotações Equipe": worksheet.column_dimensions[col_letter].width = 45
        elif col_name in ["Status/Fase", "Última Atualização", "Data da Sessão", "Dias Restantes"]: worksheet.column_dimensions[col_letter].width = 20
        elif col_name == "Valor Estimado": worksheet.column_dimensions[col_letter].width = 22
        else: worksheet.column_dimensions[col_letter].width = 16
            
    # Formatação Linha a Linha (Cores Dinâmicas)
    for row_idx in range(2, worksheet.max_row + 1):
        # Lê o status para decidir a cor da linha
        status_val = str(worksheet.cell(row=row_idx, column=status_col_idx).value).lower() if status_col_idx else ""
        dias_val = str(worksheet.cell(row=row_idx, column=dias_col_idx).value) if dias_col_idx else ""
        
        # Decide a cor de fundo (Semáforo)
        fundo_linha = PatternFill(fill_type=None)
        if row_idx % 2 == 0: fundo_linha = cor_fundo_claro
        
        if "suspen" in status_val: fundo_linha = cor_suspensa
        elif any(x in status_val for x in ["homolog", "revogad", "cancelad", "fracassad", "desert"]): fundo_linha = cor_morta

        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            col_name = worksheet.cell(row=1, column=col_idx).value
            
            cell.fill = fundo_linha
            cell.border = borda_fina
            cell.font = fonte_normal
            
            # Alertas em Vermelho para prazos curtos
            if col_name == "Dias Restantes":
                if "HOJE" in dias_val: cell.font = fonte_urgente
                elif "Faltam" in dias_val:
                    try:
                        num = int(re.search(r'\d+', dias_val).group())
                        if num <= 5: cell.font = fonte_urgente # Faltam 5 dias ou menos = Vermelho!
                    except: pass

            if col_name in ["Objeto", "Anotações Equipe"]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                texto_len = len(str(cell.value)) if cell.value else 0
                linhas_estimadas = max(1, (texto_len // 60) + 1)
                if worksheet.row_dimensions[row_idx].height is None or worksheet.row_dimensions[row_idx].height < linhas_estimadas * 15:
                    worksheet.row_dimensions[row_idx].height = linhas_estimadas * 15
            else:
                cell.alignment = Alignment(vertical="top")
                
            if col_name == "Valor Estimado" and isinstance(cell.value, (int, float)):
                cell.number_format = 'R$ #,##0.00'

    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = worksheet.dimensions

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
            
            dt_att = lic.get("dataAtualizacaoPncp")
            data_atualizacao_str = datetime.fromisoformat(dt_att[:19]).strftime('%d/%m/%Y %H:%M') if dt_att else "Não informada"

            dt_sessao = lic.get("dataAberturaProposta")
            data_sessao_str = datetime.fromisoformat(dt_sessao[:19]).strftime('%d/%m/%Y %H:%M') if dt_sessao else "Verificar Edital"

            dias_restantes = calcular_dias_restantes(data_sessao_str)

            resultados.append({
                "Identificação": identificacao,
                "Status/Fase": fase,
                "Dias Restantes": dias_restantes,
                "Data da Sessão": data_sessao_str,
                "Última Atualização": data_atualizacao_str,
                "Anotações Equipe": "",  # <--- COLUNA PARA CRM MANUAL
                "UF": uf_licitacao,
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "N/A"),
                "Modalidade": lic.get("modalidadeNome", "N/A"),
                "Objeto": objeto,
                "Valor Estimado": valor,
                "Data Publicação": lic.get("dataPublicacaoPncp")[:10] if lic.get("dataPublicacaoPncp") else "",
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
                
                df_editado = st.data_editor(
                    df_atual,
                    column_config={"Acompanhar": st.column_config.CheckboxColumn("⭐ Salvar", default=False), "Link": st.column_config.LinkColumn("Edital", display_text="Acessar 🔗")},
                    disabled=["Identificação", "Status/Fase", "Dias Restantes", "Data da Sessão", "Última Atualização", "UF", "Órgão", "Modalidade", "Objeto", "Valor Estimado", "Data Publicação", "Anotações Equipe"],
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
# ABA 2: UNIFICADOR DE INTERESSES
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
    st.markdown("Suba sua Planilha Mestre. O robô atualizará Fases, Prazos e Datas, **preservando todas as Anotações da Equipe**.")
    
    arquivo_rastreio = st.file_uploader("📂 Upload da Planilha para Rastreio (.xlsx)", type=["xlsx"], key="up_rastreio")
    
    if arquivo_rastreio is not None:
        try:
            df_rastrear = pd.read_excel(arquivo_rastreio)
            if "Link" not in df_rastrear.columns:
                st.error("❌ A planilha enviada não possui uma coluna chamada 'Link'.")
            else:
                if st.button("🔄 Rastrear e Atualizar Todos os Status"):
                    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
                    progresso = st.progress(0)
                    total = len(df_rastrear)
                    
                    if "Última Atualização" not in df_rastrear.columns: df_rastrear["Última Atualização"] = ""
                    if "Data da Sessão" not in df_rastrear.columns: df_rastrear["Data da Sessão"] = ""
                    if "Dias Restantes" not in df_rastrear.columns: df_rastrear["Dias Restantes"] = ""
                    if "Anotações Equipe" not in df_rastrear.columns: df_rastrear["Anotações Equipe"] = ""

                    with st.spinner("Consultando servidores do Governo..."):
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
                                        
                                        dt_att = dados_lic.get("dataAtualizacaoPncp")
                                        if dt_att: df_rastrear.at[index, "Última Atualização"] = datetime.fromisoformat(dt_att[:19]).strftime('%d/%m/%Y %H:%M')
                                        
                                        dt_sessao = dados_lic.get("dataAberturaProposta")
                                        if dt_sessao: 
                                            sessao_formatada = datetime.fromisoformat(dt_sessao[:19]).strftime('%d/%m/%Y %H:%M')
                                            df_rastrear.at[index, "Data da Sessão"] = sessao_formatada
                                            df_rastrear.at[index, "Dias Restantes"] = calcular_dias_restantes(sessao_formatada)
                                time.sleep(0.5) 
                            except:
                                pass
                            progresso.progress((index + 1) / total)
                    
                    st.success("✅ Rastreamento concluído! Baixe a planilha atualizada abaixo.")
                    
                    buffer_rastreio = io.BytesIO()
                    with pd.ExcelWriter(buffer_rastreio, engine='openpyxl') as writer:
                        df_rastrear.to_excel(writer, index=False, sheet_name='Base Atualizada ASA')
                        aplicar_estilo_excel(writer, df_rastrear, 'Base Atualizada ASA')
                    
                    st.download_button(
                        label="📥 Baixar Planilha Inteligente Atualizada (.xlsx)",
                        data=buffer_rastreio.getvalue(),
                        file_name=f"Master_ASA_Rastreada_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Ocorreu um erro na leitura do arquivo: {e}")
