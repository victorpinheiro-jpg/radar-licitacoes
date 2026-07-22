import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import time
import os
import io

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

# INICIALIZANDO A MEMÓRIA DO SITE
if 'licitacoes_salvas' not in st.session_state:
    st.session_state['licitacoes_salvas'] = pd.DataFrame()
if 'resultados_busca' not in st.session_state:
    st.session_state['resultados_busca'] = pd.DataFrame()
if 'busca_realizada' not in st.session_state:
    st.session_state['busca_realizada'] = False

# --- TRADUTOR DE CÓDIGOS DO GOVERNO E ESTADOS ---
MAPA_MODALIDADES = {
    "Leilão": 1, "Diálogo Competitivo": 2, "Concurso": 3, 
    "Concorrência": 4, "Pregão Eletrônico": 6
}
LISTA_UFS = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"]

# --- 2. MOTOR DE BUSCA ---
@st.cache_data(ttl=300)
def buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas):
    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
    todos_resultados = []
    erros = []
    
    if not modalidades_selecionadas:
        modalidades_selecionadas = ["Concorrência", "Leilão", "Diálogo Competitivo"]
        
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
                    else:
                        tentativas += 1; time.sleep(2) 
                except:
                    tentativas += 1; time.sleep(2)
            
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
        if estados_selecionados and (uf_licitacao not in estados_selecionados):
            continue
            
        objeto = lic.get("objetoCompra") or lic.get("sinteseObjeto") or lic.get("objeto") or "Descrição indisponível"
        objeto_str = str(objeto).lower()
        valor = lic.get("valorTotalEstimado") or 0.0 
        
        passou_palavra = any(p in objeto_str for p in lista_palavras) if (lista_palavras and lista_palavras[0] != "") else True
        if passou_palavra and (valor_min <= valor <= valor_max):
            numero_compra = lic.get("numeroCompra", "")
            ano_compra = lic.get("anoCompra", "")
            cnpj = lic.get("orgaoEntidade", {}).get("cnpj", "")
            
            if cnpj and ano_compra and numero_compra: link_final = f"https://pncp.gov.br/app/editais/{cnpj}/{ano_compra}/{numero_compra}"
            else: link_final = "https://pncp.gov.br"

            identificacao = f"{numero_compra}/{ano_compra}" if numero_compra and ano_compra else str(id_unico)

            resultados.append({
                "Identificação": identificacao,
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

aba_busca, aba_interesse = st.tabs(["🔍 Nova Busca", "⭐ Licitações de Interesse"])

with aba_busca:
    col_filtros, col_resultados = st.columns([1, 3], gap="large")

    with col_filtros:
        st.subheader("🎯 Parâmetros Regionais")
        estados_selecionados = st.multiselect("Estados de Interesse (UF):", LISTA_UFS, help="Deixe em branco para buscar no Brasil inteiro.")
        
        st.subheader("📅 Período")
        hoje = datetime.now()
        data_inicio = st.date_input("Data Inicial:", value=hoje - timedelta(days=15))
        data_fim = st.date_input("Data Final:", value=hoje)
        if data_inicio > data_fim: data_inicio, data_fim = data_fim, data_inicio
            
        st.subheader("🔎 Filtros Avançados")
        palavras_chave = st.text_input("Palavras-chave (separadas por vírgula):", "")
        modalidades_selecionadas = st.multiselect("Modalidades:", list(MAPA_MODALIDADES.keys()), default=["Concorrência", "Leilão"])
        valor_min = st.number_input("Valor Mín. (R$):", value=0.0, step=100000.0)
        valor_max = st.number_input("Valor Máx. (R$):", value=5000000000.0, step=100000.0)
        
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

        # Exibe os resultados salvos na memória, mesmo se o botão for clicado
        if st.session_state['busca_realizada']:
            df_atual = st.session_state['resultados_busca']
            if not df_atual.empty:
                if "Acompanhar" not in df_atual.columns:
                    df_atual.insert(0, "Acompanhar", False)
                
                st.write("### 📌 Resultados da Busca")
                valor_total = df_atual["Valor Estimado"].sum()
                m1, m2 = st.columns(2)
                m1.metric("Encontradas (com filtro)", f"{len(df_atual)}")
                m2.metric("Volume Financeiro", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                if "Data Publicação" in df_atual.columns:
                    df_atual = df_atual.sort_values(by="Data Publicação", ascending=False)
                
                df_editado = st.data_editor(
                    df_atual,
                    column_config={
                        "Acompanhar": st.column_config.CheckboxColumn("⭐ Salvar", default=False),
                        "Link": st.column_config.LinkColumn("Edital", display_text="Acessar 🔗")
                    },
                    disabled=["Identificação", "UF", "Órgão", "Modalidade", "Objeto", "Valor Estimado", "Data Publicação"],
                    hide_index=True, use_container_width=True,
                    key="editor_resultados"
                )
                
                if st.button("💾 Enviar selecionadas para Aba de Interesse"):
                    selecionadas = df_editado[df_editado["Acompanhar"] == True].copy()
                    selecionadas = selecionadas.drop(columns=["Acompanhar"])
                    if not selecionadas.empty:
                        st.session_state['licitacoes_salvas'] = pd.concat([st.session_state['licitacoes_salvas'], selecionadas]).drop_duplicates(subset=["Identificação"])
                        st.success("Licitações salvas com sucesso! Vá para a aba de Interesse.")
                    else:
                        st.warning("Você não marcou nenhuma licitação.")
            else:
                st.warning("Nenhuma licitação passou nos filtros. Tente apagar as palavras-chave ou remover o Estado.")

with aba_interesse:
    st.subheader("⭐ Seu Painel Temporário de Acompanhamento")
    
    if st.session_state['licitacoes_salvas'].empty:
        st.info("Você ainda não selecionou nenhuma licitação.")
    else:
        st.dataframe(
            st.session_state['licitacoes_salvas'].style.format({"Valor Estimado": "R$ {:,.2f}"}),
            column_config={"Link": st.column_config.LinkColumn("Edital", display_text="Acessar 🔗")},
            hide_index=True, use_container_width=True
        )
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state['licitacoes_salvas'].to_excel(writer, index=False, sheet_name='Oportunidades ASA')
        
        st.download_button(
            label="📥 Baixar Arquivo Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"Licitacoes_ASA_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.divider()
        if st.button("🗑️ Limpar Lista Temporária"):
            st.session_state['licitacoes_salvas'] = pd.DataFrame()
            st.rerun()
