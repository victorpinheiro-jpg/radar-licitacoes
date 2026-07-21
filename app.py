import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta

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

# --- 2. MOTOR DE BUSCA (COM RAIO-X) ---
@st.cache_data(ttl=300)
def buscar_licitacoes_periodo(data_inicio, data_fim):
    # Algumas rotas do PNCP exigem o formato YYYYMMDD
    str_inicio = data_inicio.strftime("%Y%m%d")
    str_fim = data_fim.strftime("%Y%m%d")
    
    # REDUZIMOS PARA 50 PARA NÃO TOMAR BLOQUEIO (ERRO 400)
    url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&pagina=1&tamanhoPagina=50"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            return response.json().get("data", []), 200
        else:
            # Aqui está o Raio-X: ele vai capturar o texto exato da recusa do governo
            return [], f"{response.status_code} - {response.text}"
    except Exception as e:
        return [], str(e)

def filtrar_dados(licitacoes, palavras_chave, modalidades_selecionadas):
    if not licitacoes: return pd.DataFrame()
        
    resultados = []
    lista_palavras = [p.strip().lower() for p in palavras_chave.split(',')] if palavras_chave else []
    
    for lic in licitacoes:
        objeto = str(lic.get("objeto", "")).lower()
        modalidade = lic.get("modalidadeNome", "")
        valor = lic.get("valorTotalEstimado") or 0.0 
        
        passou_palavra = any(p in objeto for p in lista_palavras) if (lista_palavras and lista_palavras[0] != "") else True
        passou_modalidade = modalidade in modalidades_selecionadas if modalidades_selecionadas else True
        
        if passou_palavra and passou_modalidade:
            resultados.append({
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "N/A"),
                "Modalidade": modalidade,
                "Objeto": lic.get("objeto"),
                "Valor Estimado": valor,
                "Link": lic.get("linkSistemaOrigem")
            })
            
    return pd.DataFrame(resultados)

# --- 3. FRONTEND ---
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
    modalidades = ["Concorrência", "Diálogo Competitivo", "Leilão", "Pregão Eletrônico"]
    modalidades_selecionadas = st.multiselect("Modalidades:", modalidades, default=[])
    
    buscar = st.button("🔍 Mapear Oportunidades")

with col_resultados:
    if buscar:
        with st.spinner("Conectando ao governo..."):
            dados_brutos, status = buscar_licitacoes_periodo(data_inicio, data_fim)
            
            if status == 200:
                st.success(f"Sucesso! O governo enviou {len(dados_brutos)} licitações neste período.")
                
                if len(dados_brutos) > 0:
                    df_final = filtrar_dados(dados_brutos, palavras_chave, modalidades_selecionadas)
                    
                    if not df_final.empty:
                        valor_total = df_final["Valor Estimado"].sum()
                        m1, m2 = st.columns(2)
                        m1.metric("Encontradas (com filtro)", f"{len(df_final)}")
                        m2.metric("Volume Financeiro", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                        
                        st.dataframe(df_final.style.format({"Valor Estimado": "R$ {:,.2f}"}), use_container_width=True)
                    else:
                        st.warning("Licitações encontradas, mas nenhuma passou nos seus filtros de Palavra/Modalidade.")
            else:
                st.error(f"O servidor do governo bloqueou a busca. Erro exato retornado por eles:")
                st.code(status) # Mostra o motivo exato do bloqueio na tela
    else:
        st.caption("👈 Configure os parâmetros ao lado e clique em Mapear Oportunidades.")
