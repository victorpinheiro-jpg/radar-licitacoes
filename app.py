import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# --- 1. MOTOR DE BUSCA (CONEXÃO COM O PNCP) ---
@st.cache_data(ttl=1800) # Guarda o resultado por 30 min para o site não travar
def buscar_licitacoes_hoje():
    # Pega a data de hoje no formato que o governo exige (AAAAMMDD)
    hoje = datetime.now().strftime("%Y%m%d")
    url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={hoje}&dataFinal={hoje}"
    
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get("data", [])
        return []
    except:
        return []

def filtrar_dados(licitacoes, palavra_chave):
    if not licitacoes:
        return pd.DataFrame()
        
    resultados = []
    termo = palavra_chave.lower()
    
    for lic in licitacoes:
        objeto = str(lic.get("objeto", "")).lower()
        
        # Se a palavra-chave estiver no objeto, ou se o campo estiver vazio, trazemos o dado
        if termo in objeto or termo == "":
            resultados.append({
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "Não informado"),
                "Objeto": lic.get("objeto"),
                "Modalidade": lic.get("modalidadeNome"),
                "Valor Estimado (R$)": lic.get("valorTotalEstimado"),
                "Link PNCP": lic.get("linkSistemaOrigem")
            })
            
    return pd.DataFrame(resultados)

# --- 2. O SITE (INTERFACE VISUAL) ---
st.set_page_config(page_title="Radar de Licitações", layout="wide")
st.title("🏗️ Monitoramento de Licitações (PNCP)")
st.caption("Acompanhamento diário com dados reais do Governo Federal")

# Filtros na Barra Lateral
st.sidebar.header("Filtros de Busca")
palavra_chave = st.sidebar.text_input("Palavra-chave no Objeto", "concessão")

# Abas de navegação
aba_novas, aba_historico = st.tabs(["🔥 Publicadas Hoje", "📊 Histórico (Em Breve)"])

with aba_novas:
    st.subheader("Buscador em Tempo Real do PNCP")
    st.write(f"Buscando licitações publicadas exatamente hoje que contenham a palavra: **{palavra_chave}**")
    
    # O botão que aciona a busca
    if st.button("Buscar Oportunidades de Hoje"):
        with st.spinner("Conectando aos servidores do governo..."):
            dados_brutos = buscar_licitacoes_hoje()
            df_filtrado = filtrar_dados(dados_brutos, palavra_chave)
            
            if not df_filtrado.empty:
                st.success(f"Sucesso! Encontramos {len(df_filtrado)} licitação(ões) hoje com a palavra '{palavra_chave}'.")
                st.dataframe(df_filtrado, use_container_width=True)
                
                # Botão para baixar o resultado real
                st.download_button(
                    label="📥 Baixar lista em Excel (CSV)",
                    data=df_filtrado.to_csv(index=False).encode('utf-8-sig'),
                    file_name=f"licitacoes_hoje_{palavra_chave}.csv",
                    mime="text/csv"
                )
            else:
                st.warning(f"Nenhuma licitação encontrada hoje com a palavra '{palavra_chave}'. Tente apagar a palavra-chave para ver todas de hoje.")

with aba_historico:
    st.info("Aqui entrará o banco de dados com o histórico de meses e anos anteriores. Faremos isso na próxima etapa!")
