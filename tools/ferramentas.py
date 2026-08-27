import os
import re
import time
import html
import json
from datetime import datetime
import requests
from ddgs import DDGS

from config import log, FUSO, DIAS, CABECALHO_HTTP

ARQUIVO_CACHE = "cache_agente.json"
# Tempo de validade do cache em segundos (86400 segundos = 24 horas)
TEMPO_EXPIRACAO_SEGUNDOS = 86400

def carregar_cache() -> dict:
    if os.path.exists(ARQUIVO_CACHE):
        try:
            with open(ARQUIVO_CACHE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"busca": {}, "pagina": {}}

def salvar_cache() -> None:
    try:
        with open(ARQUIVO_CACHE, "w", encoding="utf-8") as f:
            json.dump(_cache_geral, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning("Falha ao salvar cache no disco: %s", type(e).__name__)

def verificar_cache_valido(dicionario_cache: dict, chave: str) -> str:
    """Verifica se a chave existe no cache e se ainda está dentro da validade."""
    if chave in dicionario_cache:
        item = dicionario_cache[chave]
        if isinstance(item, str):
            return "" # Força renovação de caches antigos
            
        timestamp_salvo = item.get("timestamp", 0)
        agora = time.time()
        
        if (agora - timestamp_salvo) < TEMPO_EXPIRACAO_SEGUNDOS:
            return item.get("dado", "")
    return ""

def salvar_no_cache_memoria(dicionario_cache: dict, chave: str, valor: str) -> None:
    """Salva o dado na memória com o timestamp atual."""
    dicionario_cache[chave] = {
        "timestamp": time.time(),
        "dado": valor
    }

_cache_geral = carregar_cache()
_cache_busca = _cache_geral["busca"]
_cache_pagina = _cache_geral["pagina"]
_OPERADORES = re.compile(r'\b(site|filetype|inurl|intitle):\S+', re.IGNORECASE)


def hoje() -> str:
    """Retorna a data e a hora atuais no horário de Brasília."""
    agora = datetime.now(FUSO)
    return (f"Agora: {DIAS[agora.weekday()]}, {agora:%d/%m/%Y}, {agora:%H:%M}. "
            f"Ano corrente: {agora.year}. Mês corrente: {agora.month:02d}.")

def buscar_web(consulta: str) -> str:
    """
    Ferramenta OBRIGATÓRIA para iniciar pesquisas na web. 
    RETORNA RESUMOS. Se os resumos forem suficientes para responder a uma pergunta simples, 
    RESPONDA IMEDIATAMENTE sem abrir os sites.
    Use 'abrir_pagina' APENAS se a pergunta for complexa ou os resumos não tiverem a resposta exata.
    """
    limpa = _OPERADORES.sub("", consulta).replace('"', "").strip()
    limpa = re.sub(r"\s+", " ", limpa)
    if limpa != consulta.strip():
        log.info("   consulta sanitizada (operadores removidos)")
    if not limpa:
        return "Consulta vazia após remover operadores. Reformule sem site: ou aspas."

    chave = limpa.lower()
    dado_cache = verificar_cache_valido(_cache_busca, chave)
    if dado_cache:
        log.info("🔍 (cache válido) %s", limpa)
        return dado_cache

    log.info("🔍 buscando: %s", limpa)
    achados = []
    for tentativa in range(3):
        try:
            with DDGS(timeout=20) as ddgs:
                achados = list(ddgs.text(limpa, region="br-pt", max_results=5))
            if achados:
                break
        except Exception as e:
            if tentativa < 2:
                time.sleep(1.5 * (tentativa + 1))
                continue
            log.warning("   busca falhou: %s", type(e).__name__)
            return (f"A busca falhou ({type(e).__name__}). Não repita esta "
                    "consulta; tente termos diferentes ou informe o usuário.")

    if not achados:
        resultado = "Nenhum resultado. Não repita esta consulta; reformule ou responda com o que já tem."
        salvar_no_cache_memoria(_cache_busca, chave, resultado)
        salvar_cache()
        return resultado

    log.info("   %d resultados", len(achados))
    resultado = "\n\n".join(
        f"[{i}] {r.get('title', 'sem título')}\n{r.get('body', '')}\nURL: {r.get('href', '')}"
        for i, r in enumerate(achados, 1)
    )
    # A MENSAGEM ABAIXO FOI ALTERADA PARA REFORÇAR A ESTRATÉGIA DE PESQUISA PROGRESSIVA
    resultado += ("\n\n(Estes são apenas resumos rápidos. Se a resposta da pergunta do usuário estiver "
                  "claramente visível acima, RESPONDA AGORA. Se precisar de mais detalhes, tabelas "
                  "ou textos completos, use a ferramenta abrir_pagina em uma das URLs.)")
    
    salvar_no_cache_memoria(_cache_busca, chave, resultado)
    salvar_cache()
    return resultado

def abrir_pagina(url: str) -> str:
    """Baixa uma página web renderizando JavaScript através da API Jina Reader."""
    url = url.strip().strip('<>"\'')
    if not url.startswith(("http://", "https://")):
        return "URL inválida. Forneça um endereço iniciando com http ou https."

    FONTES_RUINS = [
        "flashscore.", "sofascore.", "cbf.com.br", "espn.com.br",
        "palmeiras.com.br", "flamengo.com.br", "corinthians.com.br",
        "saopaulofc.net", "globoesporte.globo.com/tempo-real"
    ]
    
    if any(ruim in url.lower() for ruim in FONTES_RUINS):
        log.warning("   url ignorada (fonte ruim conhecida): %s", url[:70])
        return ("Este site bloqueia acessos automatizados. IGNORE este site e TENTE A PRÓXIMA URL.")

    dado_cache = verificar_cache_valido(_cache_pagina, url)
    if dado_cache:
        log.info("📄 (cache válido) %s", url[:70])
        return dado_cache

    log.info("📄 abrindo (com Jina AI): %s", url[:70])
    
    url_jina = f"https://r.jina.ai/{url}"
    headers = {"Accept": "text/markdown"}
    
    try:
        r = requests.get(url_jina, headers=headers, timeout=25)
        r.raise_for_status()
        texto = r.text
    except Exception as e:
        log.warning("   falha ao abrir com Jina: %s", type(e).__name__)
        return f"Não foi possível abrir a página ({type(e).__name__}). Tente outra URL."

    if not texto or len(texto.strip()) < 150:
        resultado = "O texto retornado é muito curto (conteúdo vazio). IGNORE este site e tente a PRÓXIMA fonte."
        log.info("   texto muito curto ou vazio (%d chars).", len(texto if texto else ""))
    else:
        resultado = texto
        log.info("   %d caracteres extraídos e enviados na íntegra", len(texto))

    salvar_no_cache_memoria(_cache_pagina, url, resultado)
    salvar_cache()
    return resultado

def cotacao_moeda(par: str) -> str:
    """Consulta a cotação atual de moedas e criptomoedas."""
    par = par.upper().strip().replace("/", "-").replace("_", "-")
    if "-" not in par:
        par = f"{par}-BRL"

    log.info("💱 cotação: %s", par)
    try:
        r = requests.get(f"https://economia.awesomeapi.com.br/json/last/{par}", timeout=15)
        r.raise_for_status()
        d = next(iter(r.json().values()))
    except Exception as e:
        log.warning("   cotação falhou: %s", type(e).__name__)
        return f"Não foi possível obter a cotação ({type(e).__name__})."

    log.info("   %s = %s", par, d.get("bid"))
    return (
        f"{d.get('name', par)}\nCompra: {d.get('bid')}\nVenda: {d.get('ask')}\n"
        f"Variação no dia: {d.get('pctChange')}%\nMáxima: {d.get('high')} | Mínima: {d.get('low')}\n"
        f"Atualizado em: {d.get('create_date')}\nFonte: AwesomeAPI"
    )

FERRAMENTAS = [hoje, buscar_web, abrir_pagina, cotacao_moeda]
    