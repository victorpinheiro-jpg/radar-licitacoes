import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import time
import os

# --- 1. CONFIGURAÇÃO VISUAL E MEMÓRIA (PALETA PASTEL/SÓBRIA) ---
st.set_page_config(page_title="Radar de Infraestrutura | A/S", page_icon="⚖️", layout="wide")

st.markdown("""
    <style>
    /* Esconde o menu padrão e ajusta o respiro do topo */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding-top: 2rem; }
    
    /* Botões: Verde/Azul Pastel Muted (#6a9094) */
    .stButton>button { 
        width: 100%; 
        border-radius: 8px; 
        font-weight: 500; 
        background-color: #6a9094; 
        color: white; 
        height: 3em;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover { 
        background-color: #55787c; 
        color: white;
    }
    
    /* Textos, Títulos e Métricas: Tom pastel mais escuro/acinzentado para leitura suave (#436468) */
    div[data-testid="stMetricValue"] { color: #436468; font-size: 2.2rem; font-weight: 700; }
    h1, h2, h3, h4, h5, h6 { color: #436468 !important; font-weight: 600; }
    
    /* Linha divisória bem suave */
    hr { border-bottom-color: #6a9094 !important; opacity: 0.3; }
    </style>
""", unsafe_allow_html=True)

if 'licitacoes_salvas' not in st.session_state:
    st.session_state['licitacoes_salvas'] = pd.DataFrame()

# --- INSERINDO A LOGO COM ELEGÂNCIA (MENOR E CENTRALIZADA) ---
with st.sidebar:
    st.write("") # Dá um pequeno espaço no topo
    
    # Criamos 3 colunas e colocamos a logo na coluna do meio para ela ficar menor e centralizada
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists("asa_logobrasao_verde.png"):
            st.image("asa_logobrasao_verde.png", use_column_width=True)
        else:
            st.markdown("<h3 style='text-align: center; color: #436468;'>A/S</h3>", unsafe_allow_html=True)
            
    st.write("") # Dá um espaço em baixo da logo
    st.markdown("---")

# --- TRADUTOR DE CÓDIGOS DO GOVERNO ---
MAPA_MODALIDADES = {
    "Leilão": 1, 
    "Diálogo Competitivo": 2, 
    "Concurso": 3, 
    "Concorrência": 4, 
    "Pregão Eletrônico": 6
}

# --- 2. MOTOR DE BUSCA ---
@st.cache_data(ttl=300)
def buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json"
    }
    
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
            
            sucesso = False
            tentativas = 0
            max_tentativas = 3 
            
            while not sucesso and tentativas < max_tentativas:
                try:
                    response = requests.get(url, headers=headers, timeout=30)
                    if response.status_code == 200:
                        dados = response.json().get("data", [])
                        todos_resultados.extend(dados)
                        sucesso = True
                    else:
                        tentativas += 1
                        time.sleep(2) 
                except requests.exceptions.Timeout:
                    tentativas += 1
                    time.sleep(2)
                except Exception as e:
                    tentativas += 1
                    time.sleep(2)
            
            if not sucesso:
                erros.append(f"{modalidade} (bloco {str_inicio}): O Governo bloqueou após 3 tentativas.")
                
            time.sleep(1.5)
            
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
            link_original = lic.get("linkSistemaOrigem", "")
            
            if not link_original or str(link_original).strip() == "":
                link_final = None 
            elif not link_original.startswith("http"):
                link_final = "https://" + link_original
            else:
                link_final = link_original

            resultados.append({
                "Órgão": lic.get("orgaoEntidade", {}).get("razaoSocial", "N/A"),
                "Modalidade": lic.get("modalidadeNome", "N/A"),
                "Objeto": lic.get("objeto"),
                "Valor Estimado": valor,
                "Data Publicação": lic.get("dataPublicacaoPncp"),
                "Link": link_final
            })
            ids_adicionados.add(id_unico)
            
    return pd.DataFrame(resultados)

# --- 3. FRONTEND E ABAS ---
st.title("🏛️ Radar Estratégico de Licitações")
st.divider()

aba_busca, aba_interesse = st.tabs(["🔍 Nova Busca", "⭐ Licitações de Interesse"])

with aba_busca:
    col_filtros, col_resultados = st.columns([1, 3], gap="large")

    with col_filtros:
        st.subheader("🎯 Parâmetros")
        hoje = datetime.now()
        data_inicio = st.date_input("Data Inicial:", value=hoje - timedelta(days=15))
        data_fim = st.date_input("Data Final:", value=hoje)
        
        if data_inicio > data_fim:
            st.warning("⚠️ Datas invertidas corrigidas automaticamente.")
            data_inicio, data_fim = data_fim, data_inicio
            
        palavras_chave = st.text_input("Palavras-chave (vírgula):", "")
        modalidades_selecionadas = st.multiselect("Modalidades:", list(MAPA_MODALIDADES.keys()), default=["Concorrência", "Leilão"])
        
        st.subheader("💰 Filtro Financeiro")
        valor_min = st.number_input("Valor Mín. (R$):", value=0.0, step=100000.0)
        valor_max = st.number_input("Valor Máx. (R$):", value=5000000000.0, step=100000.0)
        
        buscar = st.button("🔍 Mapear Oportunidades")

    with col_resultados:
        if buscar:
            with st.spinner("Varrendo os servidores do Governo..."):
                dados_brutos, erros = buscar_licitacoes_periodo(data_inicio, data_fim, modalidades_selecionadas)
                
                if erros:
                    st.warning(f"Alguns blocos falharam (mesmo após 3 tentativas): {', '.join(erros)}")
                    
                if len(dados_brutos) > 0:
                    df_final = filtrar_dados(dados_brutos, palavras_chave, valor_min, valor_max)
                    
                    if not df_final.empty:
                        df_final.insert(0, "Acompanhar", False)
                        
                        st.write("### 📌 Resultados da Busca")
                        st.caption("Marque a caixinha 'Acompanhar' nas licitações desejadas e clique no botão abaixo da tabela.")
                        
                        valor_total = df_final["Valor Estimado"].sum()
                        m1, m2 = st.columns(2)
                        m1.metric("Encontradas (com filtro)", f"{len(df_final)}")
                        m2.metric("Volume Financeiro", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                        
                        if "Data Publicação" in df_final.columns:
                            df_final = df_final.sort_values(by="Data Publicação", ascending=False)
                        
                        df_editado = st.data_editor(
                            df_final,
                            column_config={
                                "Acompanhar": st.column_config.CheckboxColumn("⭐ Salvar", default=False),
                                "Link": st.column_config.LinkColumn("Acesso ao Edital", display_text="Acessar 🔗")
                            },
                            disabled=["Órgão", "Modalidade", "Objeto", "Valor Estimado", "Data Publicação"],
                            hide_index=True,
                            use_container_width=True
                        )
                        
                        if st.button("💾 Enviar selecionadas para Aba de Interesse"):
                            selecionadas = df_editado[df_editado["Acompanhar"] == True].copy()
                            selecionadas = selecionadas.drop(columns=["Acompanhar"])
                            
                            if not selecionadas.empty:
                                st.session_state['licitacoes_salvas'] = pd.concat([st.session_state['licitacoes_salvas'], selecionadas]).drop_duplicates(subset=["Objeto"])
                                st.success("Licitações salvas com sucesso! Vá para a aba 'Licitações de Interesse' no topo.")
                            else:
                                st.warning("Você não marcou nenhuma licitação na caixinha.")
                    else:
                        st.warning("Nenhuma licitação passou nos filtros de Valor ou Palavras-chave.")
                else:
                    st.info("Nenhuma licitação encontrada neste período ou o governo bloqueou a busca.")
        else:
            st.caption("👈 Configure os parâmetros ao lado e clique em Mapear Oportunidades.")

with aba_interesse:
    st.subheader("⭐ Seu Painel de Acompanhamento")
    st.markdown("Aqui ficam as licitações que o escritório decidiu monitorar.")
    
    if st.session_state['licitacoes_salvas'].empty:
        st.info("Você ainda não salvou nenhuma licitação de interesse.")
    else:
        st.dataframe(
            st.session_state['licitacoes_salvas'].style.format({"Valor Estimado": "R$ {:,.2f}"}),
            column_config={
                "Link": st.column_config.LinkColumn("Edital", display_text="Acessar 🔗")
            },
            hide_index=True,
            use_container_width=True
        )
        
        st.download_button(
            label="📥 Exportar Lista de Interesse para o Excel",
            data=st.session_state['licitacoes_salvas'].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig'),
            file_name="licitacoes_interesse.csv",
            mime="text/csv"
        )
        
        if st.button("🗑️ Limpar Lista de Interesse"):
            st.session_state['licitacoes_salvas'] = pd.DataFrame()
            st.rerun()()
            st.rerun()
