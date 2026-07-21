import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import time

# --- 1. CONFIGURAÇÃO VISUAL PREMIUM ---
st.set_page_config(page_title="Radar de Infraestrutura", page_icon="⚖️", layout="wide")

st.markdown("""
    <style>
    .stButton>button {
        width: 100%; background-color: #1E3A8A; color: white;
        border-radius: 8px; height: 3em; font-weight: bold;
    }
    .stButton>button:hover { background-color: #1e40af; }
    div[data-testid="stMetricValue"] { color: #1E3A8A; font-size: 2rem; }
    </style>
""", unsafe_allow_html=True)

# --- TRADUTOR DE CÓDIGOS DO GOVERNO ---
# O Governo agora exige o código numérico exato de cada modalidade
MAPA_MODALIDADES = {
    "Leilão": 1,
    "Diálogo Competitivo": 2,
    "Concurso": 3,
    "Concorrência": 4,
    "Pregão Eletrônico": 6
}

# --- 2. MOTOR DE BUSCA INTELIGENTE ---
@st.cache_data(ttl=300)
def buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas):
    str_inicio = data_inicio.strftime("%Y%m%d")
    str_fim = data_fim.strftime("%Y%m%d")
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json"
    }
    
    todos_resultados = []
    erros = []
    
    # Se você deixar vazio, o robô busca as principais por padrão
    if not modalidades_selecionadas:
        modalidades_selecionadas = ["Concorrência", "Leilão", "Diálogo Competitivo", "Pregão Eletrônico"]
        
    for modalidade in modalidades_selecionadas:
        codigo = MAPA_MODALIDADES.get(modalidade)
        if not codigo:
            continue
            
        # Agora estamos enviando a chave que o governo exigiu (&codigoModalidadeContratacao=X)
        url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&codigoModalidadeContratacao={codigo}&pagina=1&tamanhoPagina=50"
        
        try:
            response = requests.get(url, headers=headers, timeout=15)
            if response.status_code == 200:
                dados = response.json().get("data", [])
                todos_resultados.extend(dados)
            else:
                erros.append(f"{modalidade}: Erro {response.status_code}")
        except Exception as e:
            erros.append(f"{modalidade}: Falha na conexão")
            
        # Pequena pausa para o governo não achar que somos um ataque hacker (DDoS)
        time.sleep(0.5)
        
    return todos_resultados, erros

def filtrar_palavras(licitacoes, palavras_chave):
    if not licitacoes: return pd.DataFrame()
        
    resultados = []
    lista_palavras = [p.strip().lower() for p in palavras_chave.split(',')] if palavras_chave else []
    
    for lic in licitacoes:
        objeto = str(lic.get("objeto", "")).lower()
        valor = lic.get("valorTotalEstimado") or 0.0 
        
        passou_palavra = any(p in objeto for p in lista_palavras) if (lista_palavras and lista_palavras[0] != "") else True
        
        if passou_palavra:
            resultados.append({
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "N/A"),
                "Modalidade": lic.get("modalidadeNome", "N/A"),
                "Objeto": lic.get("objeto"),
                "Valor Estimado": valor,
                "Link": lic.get("linkSistemaOrigem")
            })
            
    return pd.DataFrame(resultados)

# --- 3. FRONTEND (A TELA DO SITE) ---
st.title("🏛️ Radar Estratégico de Licitações")
st.markdown("Monitoramento inteligente via **Portal Nacional de Contratações Públicas (PNCP)**.")
st.divider()

col_filtros, col_resultados = st.columns([1, 3], gap="large")

with col_filtros:
    st.subheader("🎯 Parâmetros")
    
    hoje = datetime.now()
    data_inicio = st.date_input("De (Data Inicial):", value=hoje)
    data_fim = st.date_input("Até (Data Final):", value=hoje)
    
    palavras_chave = st.text_input("Palavras-chave (vírgula):", "")
    
    modalidades_selecionadas = st.multiselect(
        "Modalidades:", 
        list(MAPA_MODALIDADES.keys()), 
        default=["Concorrência", "Leilão"]
    )
    
    buscar = st.button("🔍 Mapear Oportunidades")

with col_resultados:
    if buscar:
        with st.spinner("Buscando dados nos servidores do Governo..."):
            dados_brutos, erros = buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas)
            
            if erros:
                st.warning(f"Alguns alertas durante a busca: {', '.join(erros)}")
                
            if len(dados_brutos) > 0:
                st.success(f"Sucesso! O governo liberou {len(dados_brutos)} licitações brutas nas modalidades selecionadas.")
                
                df_final = filtrar_palavras(dados_brutos, palavras_chave)
                
                if not df_final.empty:
                    valor_total = df_final["Valor Estimado"].sum()
                    m1, m2 = st.columns(2)
                    m1.metric("Encontradas (com filtro)", f"{len(df_final)}")
                    m2.metric("Volume Financeiro", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    
                    st.dataframe(df_final.style.format({"Valor Estimado": "R$ {:,.2f}"}), use_container_width=True)
                else:
                    st.warning("O governo retornou dados, mas nenhuma passou nos seus filtros de Palavras-chave. Tente apagar a palavra-chave.")
            else:
                if not erros:
                    st.info("Nenhuma licitação publicada neste período para estas modalidades.")
    else:
        st.caption("👈 Configure os parâmetros ao lado e clique em Mapear Oportunidades.")
