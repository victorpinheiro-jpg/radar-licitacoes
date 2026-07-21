import streamlit as st
import pandas as pd

# 1. Configuração da Página
st.set_page_config(page_title="Radar de Licitações - Infraestrutura", layout="wide")

st.title("🏗️ Monitoramento de Licitações de Infraestrutura (PNCP)")
st.caption("Acompanhamento estratégico para prospecção e recursos contratuais")

# 2. Barra Lateral para Filtros
st.sidebar.header("Filtros de Busca")
modalidade = st.sidebar.multiselect(
    "Modalidade",
    ["Concorrência", "Diálogo Competitivo", "Leilão", "Pregão Eletrônico"],
    default=["Concorrência", "Diálogo Competitivo"]
)
termo_busca = st.sidebar.text_input("Palavra-chave no Objeto", "Porto")

# 3. Criando as Abas do Site
aba_novas, aba_historico, aba_gestao = st.tabs(["🔥 Novas Hoje", "📊 Histórico (Desde 2024)", "💼 Gestão de Clientes"])

with aba_novas:
    st.subheader("Oportunidades Publicadas Hoje")
    st.info("Buscando atualizações mais recentes no PNCP...")

with aba_historico:
    st.subheader("Base Geral de Editais")
    
    # Dados simulados para você ver o visual da tabela
    dados_exemplo = pd.DataFrame({
        "Órgão": ["ANTAQ", "EPL / Infra S.A.", "Gov. Estado SP"],
        "Objeto": ["Arrendamento de Terminal Portuário no Porto de Santos", "Concessão de Rodovia BR-101", "Parceria Público-Privada Trem Intercidades"],
        "Modalidade": ["Leilão", "Concorrência", "Diálogo Competitivo"],
        "Valor (R$)": [250000000.00, 1200000000.00, 5000000000.00],
        "Status": ["Aguardando Propostas", "Em Julgamento", "Edital Publicado"]
    })
    
    st.dataframe(dados_exemplo, use_container_width=True)
    
    st.download_button(
        label="📥 Baixar esta lista no Excel",
        data=dados_exemplo.to_csv(index=False),
        file_name="licitacoes_infraestrutura.csv",
        mime="text/csv"
    )

with aba_gestao:
    st.subheader("Acompanhamento de Prospecção")
    st.write("Em breve: Marque quais clientes o escritório está contatando para cada edital.")
