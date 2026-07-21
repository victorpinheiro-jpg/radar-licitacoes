import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta

# --- 1. MOTOR DE BUSCA AVANÇADO ---
@st.cache_data(ttl=1800) 
def buscar_licitacoes_periodo(data_inicio, data_fim):
    # Converte as datas do calendário para o formato do governo
    str_inicio = data_inicio.strftime("%Y%m%d")
    str_fim = data_fim.strftime("%Y%m%d")
    
    # Busca até 500 resultados por vez (limite do sistema deles)
    url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&pagina=1&tamanhoPagina=500"
    
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get("data", [])
        return []
    except:
        return []

def filtrar_dados(licitacoes, palavras_chave, modalidades_selecionadas, valor_min, valor_max):
    if not licitacoes:
        return pd.DataFrame()
        
    resultados = []
    
    # Transforma o texto digitado em uma lista de palavras (separadas por vírgula)
    lista_palavras = [p.strip().lower() for p in palavras_chave.split(',')] if palavras_chave else []
    
    for lic in licitacoes:
        objeto = str(lic.get("objeto", "")).lower()
        modalidade_atual = lic.get("modalidadeNome", "")
        # Se não tiver valor informado, assume 0 para não quebrar a conta
        valor_estimado = lic.get("valorTotalEstimado") or 0.0 
        
        # 1. Filtra Palavras-chave
        passou_palavra = True
        if lista_palavras and lista_palavras[0] != "":
            passou_palavra = any(palavra in objeto for palavra in lista_palavras)
            
        # 2. Filtra Modalidade
        passou_modalidade = True
        if modalidades_selecionadas:
            passou_modalidade = modalidade_atual in modalidades_selecionadas
            
        # 3. Filtra Valores
        passou_valor = (valor_min <= valor_estimado <= valor_max)
        
        # Se passar por todos os filtros, salva para o escritório
        if passou_palavra and passou_modalidade and passou_valor:
            resultados.append({
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "Não informado"),
                "Objeto": lic.get("objeto"),
                "Modalidade": modalidade_atual,
                "Valor Estimado (R$)": valor_estimado,
                "Data Publicação": lic.get("dataPublicacaoPncp"),
                "Link PNCP": lic.get("linkSistemaOrigem")
            })
            
    return pd.DataFrame(resultados)

# --- 2. O SITE (INTERFACE VISUAL) ---
st.set_page_config(page_title="Radar de Infraestrutura", layout="wide")
st.title("🏗️ Radar Estratégico de Licitações (PNCP)")

# BARRA LATERAL (OS FILTROS REAIS AGORA)
st.sidebar.header("🎯 Parâmetros de Busca")

hoje = datetime.now()
data_inicio = st.sidebar.date_input("Data Inicial", hoje - timedelta(days=7))
data_fim = st.sidebar.date_input("Data Final", hoje)

palavras_chave = st.sidebar.text_input("Palavras-chave (separe por vírgula)", "concessão, rodovia, porto, ferrovia, ppp")

modalidades_disponiveis = [
    "Concorrência", "Diálogo Competitivo", "Leilão", 
    "Pregão Eletrônico", "Concurso", "Dispensa de Licitação", "Inexigibilidade"
]
modalidades_selecionadas = st.sidebar.multiselect(
    "Modalidades", 
    modalidades_disponiveis, 
    default=["Concorrência", "Diálogo Competitivo", "Leilão"]
)

st.sidebar.subheader("💰 Faixa de Valor (R$)")
valor_min = st.sidebar.number_input("Valor Mínimo (R$)", value=0.0, step=100000.0, format="%f")
valor_max = st.sidebar.number_input("Valor Máximo (R$)", value=5000000000.0, step=100000.0, format="%f")

# TELA PRINCIPAL
st.write("Ajuste os filtros na barra lateral e clique no botão abaixo para extrair os dados do portal do governo.")

if st.button("Buscar Oportunidades", type="primary"):
    with st.spinner("Varrendo a base de dados do Governo Federal..."):
        # 1. Puxa tudo no período
        dados_brutos = buscar_licitacoes_periodo(data_inicio, data_fim)
        
        # 2. Passa o pente fino dos filtros do escritório
        df_filtrado = filtrar_dados(dados_brutos, palavras_chave, modalidades_selecionadas, valor_min, valor_max)
        
        if not df_filtrado.empty:
            st.success(f"Bingo! Encontramos {len(df_filtrado)} licitação(ões) nos critérios exatos da busca.")
            
            # Formata a tabela para ficar elegante
            st.dataframe(
                df_filtrado.style.format({"Valor Estimado (R$)": "R$ {:,.2f}"}), 
                use_container_width=True
            )
            
            st.download_button(
                label="📥 Exportar relatório para o Excel (CSV)",
                data=df_filtrado.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig'),
                file_name="relatorio_prospeccao.csv",
                mime="text/csv"
            )
        else:
            st.warning("Nenhuma licitação bateu com todos esses filtros simultaneamente. Tente ampliar as datas ou remover algumas palavras-chave.")
