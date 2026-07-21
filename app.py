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
MAPA_MODALIDADES = {
    "Leilão": 1,
    "Diálogo Competitivo": 2,
    "Concurso": 3,
    "Concorrência": 4,
    "Pregão Eletrônico": 6
}

# --- 2. MOTOR DE BUSCA (COM FATIADOR DE DATAS) ---
@st.cache_data(ttl=300)
def buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json"
    }
    
    todos_resultados = []
    erros = []
    
    if not modalidades_selecionadas:
        modalidades_selecionadas = ["Concorrência", "Leilão", "Diálogo Competitivo"]
        
    # FATIADOR: Quebra o período grande em pedaços de 30 dias
    chunks = []
    atual = data_inicio
    while atual <= data_fim:
        proximo = atual + timedelta(days=30)
        if proximo > data_fim:
            proximo = data_fim
        chunks.append((atual, proximo))
        atual = proximo + timedelta(days=1)
        
    for modalidade in modalidades_selecionadas:
        codigo = MAPA_MODALIDADES.get(modalidade)
        if not codigo: continue
            
        for inicio_chunk, fim_chunk in chunks:
            str_inicio = inicio_chunk.strftime("%Y%m%d")
            str_fim = fim_chunk.strftime("%Y%m%d")
            
            url = f"https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao?dataInicial={str_inicio}&dataFinal={str_fim}&codigoModalidadeContratacao={codigo}&pagina=1&tamanhoPagina=50"
            
            try:
                response = requests.get(url, headers=headers, timeout=30)
                if response.status_code == 200:
                    dados = response.json().get("data", [])
                    todos_resultados.extend(dados)
                else:
                    erros.append(f"{modalidade} (bloco {str_inicio}): Erro {response.status_code}")
            except requests.exceptions.Timeout:
                erros.append(f"{modalidade} (bloco {str_inicio}): O Governo demorou muito a responder (Timeout)")
            except Exception as e:
                erros.append(f"{modalidade} (bloco {str_inicio}): Falha de conexão")
                
            time.sleep(1)
            
    return todos_resultados, list(set(erros))

def filtrar_dados(licitacoes, palavras_chave, valor_min, valor_max):
    if not licitacoes: return pd.DataFrame()
        
    resultados = []
    lista_palavras = [p.strip().lower() for p in palavras_chave.split(',')] if palavras_chave else []
    ids_adicionados = set()
    
    for lic in licitacoes:
        id_unico = lic.get("id") or lic.get("linkSistemaOrigem")
        if id_unico in ids_adicionados:
            continue
            
        objeto = str(lic.get("objeto", "")).lower()
        valor = lic.get("valorTotalEstimado") or 0.0 
        
        passou_palavra = any(p in objeto for p in lista_palavras) if (lista_palavras and lista_palavras[0] != "") else True
        passou_valor = (valor_min <= valor <= valor_max)
        
        if passou_palavra and passou_valor:
            # Pegando o link com segurança, caso venha vazio da API do governo
            link_original = lic.get("linkSistemaOrigem", "")
            
            resultados.append({
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "N/A"),
                "Modalidade": lic.get("modalidadeNome", "N/A"),
                "Objeto": lic.get("objeto"),
                "Valor Estimado": valor,
                "Data Publicação": lic.get("dataPublicacaoPncp"),
                "Link": link_original if link_original else "Sem Link"
            })
            ids_adicionados.add(id_unico)
            
    return pd.DataFrame(resultados)

# --- 3. FRONTEND (A TELA DO SITE) ---
st.title("🏛️ Radar Estratégico de Licitações")
st.markdown("Monitoramento inteligente via **Portal Nacional de Contratações Públicas (PNCP)**.")
st.divider()

col_filtros, col_resultados = st.columns([1, 3], gap="large")

with col_filtros:
    st.subheader("🎯 Parâmetros")
    
    hoje = datetime.now()
    data_inicio = st.date_input("De (Data Inicial):", value=hoje - timedelta(days=30))
    data_fim = st.date_input("Até (Data Final):", value=hoje)
    
    if data_inicio > data_fim:
        st.warning("⚠️ Detectamos datas invertidas. O sistema ajustou a ordem automaticamente para buscar!")
        data_inicio, data_fim = data_fim, data_inicio
    
    palavras_chave = st.text_input("Palavras-chave (vírgula):", "")
    
    modalidades_selecionadas = st.multiselect(
        "Modalidades:", 
        list(MAPA_MODALIDADES.keys()), 
        default=["Concorrência", "Leilão"]
    )
    
    st.subheader("💰 Filtro Financeiro")
    valor_min = st.number_input("Valor Mínimo (R$):", min_value=0.0, value=0.0, step=100000.0)
    valor_max = st.number_input("Valor Máximo (R$):", min_value=0.0, value=5000000000.0, step=100000.0)
    
    buscar = st.button("🔍 Mapear Oportunidades")

with col_resultados:
    if buscar:
        with st.spinner("Varrendo os servidores do Governo (pode levar alguns segundos em períodos longos)..."):
            dados_brutos, erros = buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas)
            
            if erros:
                st.warning(f"Alguns blocos de data falharam devido a instabilidade do governo: {', '.join(erros)}")
                
            if len(dados_brutos) > 0:
                df_final = filtrar_dados(dados_brutos, palavras_chave, valor_min, valor_max)
                
                if not df_final.empty:
                    st.success(f"Sucesso! Encontramos {len(df_final)} licitação(ões) dentro dos seus critérios.")
                    
                    valor_total = df_final["Valor Estimado"].sum()
                    m1, m2 = st.columns(2)
                    m1.metric("Encontradas (com filtro)", f"{len(df_final)}")
                    m2.metric("Volume Financeiro", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    
                    if "Data Publicação" in df_final.columns:
                        df_final = df_final.sort_values(by="Data Publicação", ascending=False)
                        
                    # Configuração para transformar o texto da coluna 'Link' em um botão clicável
                    st.dataframe(
                        df_final.style.format({"Valor Estimado": "R$ {:,.2f}"}), 
                        use_container_width=True,
                        column_config={
                            "Link": st.column_config.LinkColumn(
                                "Acesso ao Edital",
                                display_text="Acessar Edital 🔗"
                            )
                        }
                    )
                else:
                    st.warning("O governo retornou dados, mas nenhuma licitação bateu com os seus filtros de Valor ou Palavras-chave.")
            else:
                if not erros:
                    st.info("Nenhuma licitação publicada neste período para estas modalidades.")
    else:
        st.caption("👈 Configure os parâmetros ao lado e clique em Mapear Oportunidades.")
