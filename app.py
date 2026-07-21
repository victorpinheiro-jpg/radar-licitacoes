import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta

# --- 1. CONFIGURAÇÃO VISUAL ---
st.set_page_config(page_title="Radar de Infraestrutura", page_icon="⚖️", layout="wide")

st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        background-color: #1E3A8A;
        color: white;
        border-radius: 8px;
        height: 3em;
        font-weight: bold;
    }
    .stButton>button:hover { background-color: #1e40af; }
    div[data-testid="stMetricValue"] { color: #1E3A8A; font-size: 2rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. MOTOR DE BUSCA (COM DIAGNÓSTICO) ---
@st.cache_data(ttl=300) # Reduzi o cache para 5 minutos para facilitar testes
def buscar_licitacoes_periodo(data_inicio, data_fim):
    str_inicio = data_inicio.strftime("%Y%m%d")
    str_fim = data_fim.strftime("%Y%m%d")
    
    url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&pagina=1&tamanhoPagina=500"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        # Modo diagnóstico: Retorna também o código de status para sabermos o que houve
        if response.status_code == 200:
            return response.json().get("data", []), 200
        else:
            return [], response.status_code
    except Exception as e:
        return [], str(e)

def filtrar_dados(licitacoes, palavras_chave, modalidades_selecionadas):
    if not licitacoes:
        return pd.DataFrame()
        
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
                "Link do Edital": lic.get("linkSistemaOrigem")
            })
            
    return pd.DataFrame(resultados)

# --- 3. FRONTEND ---
st.title("🏛️ Radar Estratégico de Licitações")
st.markdown("Monitoramento inteligente de editais e concessões via **Portal Nacional de Contratações Públicas (PNCP)**.")
st.divider()

col_filtros, col_resultados = st.columns([1, 3], gap="large")

with col_filtros:
    st.subheader("🎯 Parâmetros")
    
    # Datas padrão sugeridas: últimos 7 dias
    hoje = datetime.now()
    data_inicio = st.date_input("De (Data Inicial):", value=hoje - timedelta(days=7))
    data_fim = st.date_input("Até (Data Final):", value=hoje)
    
    # Deixei a palavra-chave em branco por padrão para forçar o sistema a trazer tudo primeiro
    palavras_chave = st.text_input("Palavras-chave (separadas por vírgula):", "")
    
    # Deixei as modalidades vazias por padrão para não restringir a busca inicial
    modalidades = ["Concorrência", "Diálogo Competitivo", "Leilão", "Pregão Eletrônico"]
    modalidades_selecionadas = st.multiselect("Modalidades:", modalidades, default=[])
    
    buscar = st.button("🔍 Mapear Oportunidades")

with col_resultados:
    if buscar:
        with st.spinner("Conectando ao banco de dados do Governo..."):
            dados_brutos, status = buscar_licitacoes_periodo(data_inicio, data_fim)
            
            # --- PAINEL DE DIAGNÓSTICO ---
            if status == 200:
                st.success(f"Conexão bem-sucedida! O governo nos enviou {len(dados_brutos)} licitações neste período.")
            elif isinstance(status, int):
                st.error(f"O servidor do governo bloqueou a busca. Código de Erro: {status}.")
                st.info("💡 Erro 403 significa que o governo bloqueou o endereço de IP do nosso site por segurança. Erro 500 significa que o site deles caiu.")
            else:
                st.error(f"Erro de conexão com a internet: {status}")
            
            # Só filtra se tiver dados
            if isinstance(status, int) and status == 200 and len(dados_brutos) > 0:
                df_final = filtrar_dados(dados_brutos, palavras_chave, modalidades_selecionadas)
                
                if not df_final.empty:
                    st.info(f"Após aplicar seus filtros de palavra-chave e modalidade, restaram {len(df_final)} oportunidades.")
                    
                    valor_total = df_final["Valor Estimado"].sum()
                    m1, m2 = st.columns(2)
                    m1.metric("Oportunidades Encontradas", f"{len(df_final)} editais")
                    m2.metric("Volume Financeiro Total", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    
                    st.dataframe(df_final.style.format({"Valor Estimado": "R$ {:,.2f}"}), use_container_width=True, height=400)
                else:
                    st.warning("A busca retornou dados do governo, mas NENHUMA licitação passou nos seus filtros. Tente apagar as palavras-chave e limpar as modalidades para ver a lista completa.")
    else:
        st.caption("👈 Configure os parâmetros ao lado e clique em Mapear Oportunidades. Para o primeiro teste, recomendo deixar 'Palavras-chave' e 'Modalidades' em branco!")
