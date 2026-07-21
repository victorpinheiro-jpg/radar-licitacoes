import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# --- 1. CONFIGURAÇÃO VISUAL PREMIUM ---
st.set_page_config(page_title="Radar de Infraestrutura", page_icon="⚖️", layout="wide")

# CSS Customizado para deixar o site mais bonito
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
    .stButton>button:hover {
        background-color: #1e40af;
        border-color: #1e40af;
    }
    div[data-testid="stMetricValue"] {
        font-size: 2rem;
        color: #1E3A8A;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. MOTOR DE BUSCA (API PNCP) ---
@st.cache_data(ttl=1800) 
def buscar_licitacoes_periodo(data_inicio, data_fim):
    str_inicio = data_inicio.strftime("%Y%m%d")
    str_fim = data_fim.strftime("%Y%m%d")
    
    url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&pagina=1&tamanhoPagina=500"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get("data", [])
        return []
    except:
        return []

def filtrar_dados(licitacoes, palavras_chave, modalidades_selecionadas):
    if not licitacoes:
        return pd.DataFrame()
        
    resultados = []
    lista_palavras = [p.strip().lower() for p in palavras_chave.split(',')] if palavras_chave else []
    
    for lic in licitacoes:
        objeto = str(lic.get("objeto", "")).lower()
        modalidade = lic.get("modalidadeNome", "")
        valor = lic.get("valorTotalEstimado") or 0.0 
        
        # Filtros
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

# --- 3. CONSTRUÇÃO DO PAINEL (FRONTEND) ---
st.title("🏛️ Radar Estratégico de Licitações")
st.markdown("Monitoramento inteligente de editais e concessões via **Portal Nacional de Contratações Públicas (PNCP)**.")
st.divider()

# Dividindo a tela em Duas Colunas (Filtros na esquerda, Resultados na direita)
col_filtros, col_resultados = st.columns([1, 3], gap="large")

with col_filtros:
    st.subheader("🎯 Parâmetros")
    
    data_inicio = st.date_input("De (Data Inicial):", value=datetime(2023, 10, 1))
    data_fim = st.date_input("Até (Data Final):", value=datetime(2023, 10, 30))
    
    palavras_chave = st.text_input("Palavras-chave (separadas por vírgula):", "concessão, rodovia, porto")
    
    modalidades = ["Concorrência", "Diálogo Competitivo", "Leilão", "Pregão Eletrônico"]
    modalidades_selecionadas = st.multiselect("Modalidades:", modalidades, default=["Concorrência"])
    
    buscar = st.button("🔍 Mapear Oportunidades")

with col_resultados:
    if buscar:
        with st.spinner("Conectando ao banco de dados do Governo..."):
            dados_brutos = buscar_licitacoes_periodo(data_inicio, data_fim)
            df_final = filtrar_dados(dados_brutos, palavras_chave, modalidades_selecionadas)
            
            if not df_final.empty:
                # Criação dos Cards de Métricas
                valor_total = df_final["Valor Estimado"].sum()
                
                m1, m2 = st.columns(2)
                m1.metric("Oportunidades Encontradas", f"{len(df_final)} editais")
                m2.metric("Volume Financeiro Total", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                st.write("### Lista de Processos")
                
                # Formatação elegante da tabela
                st.dataframe(
                    df_final.style.format({"Valor Estimado": "R$ {:,.2f}"}),
                    use_container_width=True,
                    height=400
                )
                
                st.download_button(
                    label="📥 Exportar Dados para Excel",
                    data=df_final.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig'),
                    file_name="prospeccao_escritorio.csv",
                    mime="text/csv"
                )
            else:
                st.info("Nenhuma licitação encontrada com estes critérios. Tente remover algumas palavras-chave ou alterar as datas para o ano de 2023/2024.")
    else:
        # Tela inicial de espera
        st.caption("👈 Configure os parâmetros ao lado e clique em Mapear Oportunidades para iniciar.")
