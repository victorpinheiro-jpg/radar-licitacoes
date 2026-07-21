# --- 2. MOTOR DE BUSCA (COM INSISTÊNCIA AUTOMÁTICA) ---
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
        
    # Quebra o período grande em pedaços de 15 dias (mais leve para o governo)
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
            max_tentativas = 3 # O robô vai tentar 3 vezes antes de desistir
            
            while not sucesso and tentativas < max_tentativas:
                try:
                    response = requests.get(url, headers=headers, timeout=30)
                    if response.status_code == 200:
                        dados = response.json().get("data", [])
                        todos_resultados.extend(dados)
                        sucesso = True
                    else:
                        tentativas += 1
                        time.sleep(2) # Espera 2 segs se der erro antes de tentar de novo
                except requests.exceptions.Timeout:
                    tentativas += 1
                    time.sleep(2)
                except Exception as e:
                    tentativas += 1
                    time.sleep(2)
            
            # Se tentou 3 vezes e não conseguiu, aí sim avisa o usuário
            if not sucesso:
                erros.append(f"{modalidade} (bloco {str_inicio}): O Governo bloqueou após 3 tentativas.")
                
            # Pausa de 1.5 segundos entre cada bloco que deu certo
            time.sleep(1.5)
            
    return todos_resultados, list(set(erros))
