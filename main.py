"""
main.py
Agente autônomo de pesquisa com Gemini API.

Ferramentas:
- buscar_web     → DuckDuckGo/ddgs, retorna snippets curtos
- abrir_pagina   → baixa e extrai o texto de uma URL (tabelas, calendários)
- cotacao_moeda  → AwesomeAPI, dado numérico em tempo real
- hoje           → data e hora atuais para ancoragem temporal

Requisitos:
    pip install google-genai python-dotenv ddgs requests tzdata
.env:
    GEMINI_API_KEY=sua_chave
"""

import os
import re
import sys
import time
import html
import random
import logging
from datetime import datetime, timedelta, timezone
from dataclasses import dataclass, field

import requests
from dotenv import load_dotenv
from ddgs import DDGS
from google import genai
from google.genai import types
from google.genai import errors as genai_errors

# ---------------------------------------------------------------- CONFIGURAÇÃO
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
)
for ruidoso in ("google_genai", "httpx", "httpcore", "ddgs",
                "ddgs.engines", "primp", "urllib3", "requests"):
    logging.getLogger(ruidoso).setLevel(logging.ERROR)
logging.getLogger("ddgs").propagate = False
log = logging.getLogger("agente")

try:
    from zoneinfo import ZoneInfo
    FUSO = ZoneInfo("America/Sao_Paulo")
except Exception:
    FUSO = timezone(timedelta(hours=-3), name="BRT")
    log.warning("tzdata ausente — usando UTC-3 fixo. Instale com: pip install tzdata")

DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
        "sexta-feira", "sábado", "domingo"]

CABECALHO_HTTP = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/126.0 Safari/537.36"),
    "Accept-Language": "pt-BR,pt;q=0.9",
}

CANDIDATOS = [
    "gemini-3.5-flash-lite",
    "gemini-3.6-flash",
    "gemini-flash-latest",
]


def contexto_temporal() -> str:
    agora = datetime.now(FUSO)
    return (f"{DIAS[agora.weekday()]}, {agora:%d/%m/%Y} às {agora:%H:%M} "
            f"(horário de Brasília)")


def instrucao_com_busca() -> str:
    return (
        "Você é um agente de pesquisa objetivo. Responda em português do Brasil.\n\n"
        f"DATA E HORA ATUAIS: {contexto_temporal()}.\n\n"
        "FERRAMENTAS:\n"
        "- hoje: confirme a data antes de qualquer raciocínio temporal.\n"
        "- cotacao_moeda: use SEMPRE para dólar, euro, bitcoin e câmbio. "
        "Jamais responda cotação a partir de buscar_web.\n"
        "- buscar_web: retorna apenas RESUMOS curtos. Serve para descobrir "
        "quais páginas existem, não para extrair listas ou tabelas.\n"
        "- abrir_pagina: baixa o conteúdo real de uma URL. Use SEMPRE que a "
        "pergunta envolver calendários, tabelas, agendas, listas de jogos, "
        "escalações, preços ou qualquer dado que não caiba em um resumo.\n\n"
        "FLUXO OBRIGATÓRIO para perguntas do tipo 'próximo jogo', 'agenda', "
        "'quando é', 'tabela':\n"
        "1. buscar_web para localizar a página de agenda oficial.\n"
        "2. abrir_pagina na URL mais promissora para ler a tabela inteira.\n"
        "3. Liste mentalmente TODAS as datas futuras encontradas.\n"
        "4. Escolha a MENOR data que seja igual ou posterior a hoje.\n"
        "Nunca declare 'o próximo jogo é X' baseado só em snippets.\n\n"
        "REGRAS DE BUSCA:\n"
        "- Nunca use operadores como site:, filetype: ou aspas de frase exata; "
        "os backends rejeitam e a busca falha.\n"
        "- Use o ano corrente, nunca 'ano A ou ano B'.\n"
        "- Máximo de 3 buscas e 2 aberturas de página por pergunta.\n"
        "- Não repita consultas quase idênticas.\n\n"
        "HONESTIDADE:\n"
        "- Cite os links usados.\n"
        "- Se não conseguiu abrir nenhuma página, diga que a resposta vem de "
        "resumos e pode estar incompleta.\n"
        "- Se o usuário apontar um erro seu, explique a causa real com base no "
        "que as ferramentas retornaram. Não invente justificativas técnicas.\n"
        "- Nunca invente dados."
    )


def instrucao_sem_busca() -> str:
    return (
        "Você é um agente de pesquisa objetivo. Responda em português do Brasil. "
        f"Data atual: {contexto_temporal()}. "
        "Você NÃO tem acesso à internet nesta sessão. Se a pergunta exigir dados "
        "atuais, diga isso claramente em vez de inventar."
    )


# ----------------------------------------------------------------- FERRAMENTAS
_cache_busca: dict[str, str] = {}
_cache_pagina: dict[str, str] = {}

_OPERADORES = re.compile(r'\b(site|filetype|inurl|intitle):\S+', re.IGNORECASE)


def hoje() -> str:
    """Retorna a data e a hora atuais no horário de Brasília.

    Use antes de qualquer raciocínio que envolva 'hoje', 'amanhã',
    'próximo', 'este mês' ou comparação de datas.

    Returns:
        Dia da semana, data e hora atuais.
    """
    agora = datetime.now(FUSO)
    return (f"Agora: {DIAS[agora.weekday()]}, {agora:%d/%m/%Y}, {agora:%H:%M}. "
            f"Ano corrente: {agora.year}. Mês corrente: {agora.month:02d}.")


def buscar_web(consulta: str) -> str:
    """Busca páginas na internet e retorna resumos curtos.

    Os resumos têm poucas linhas e NÃO contêm tabelas ou listas completas.
    Depois de encontrar a página certa, use abrir_pagina para ler o conteúdo.
    Não use operadores como site: ou aspas de frase exata.

    Args:
        consulta: termos de busca em linguagem natural, com o ano correto.

    Returns:
        Até 5 resultados com título, resumo e link.
    """
    limpa = _OPERADORES.sub("", consulta).replace('"', "").strip()
    limpa = re.sub(r"\s+", " ", limpa)
    if limpa != consulta.strip():
        log.info("   consulta sanitizada (operadores removidos)")
    if not limpa:
        return "Consulta vazia após remover operadores. Reformule sem site: ou aspas."

    chave = limpa.lower()
    if chave in _cache_busca:
        log.info("🔍 (cache) %s", limpa)
        return _cache_busca[chave]

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
        resultado = ("Nenhum resultado. Não repita esta consulta; "
                     "reformule ou responda com o que já tem.")
        _cache_busca[chave] = resultado
        return resultado

    log.info("   %d resultados", len(achados))
    resultado = "\n\n".join(
        f"[{i}] {r.get('title', 'sem título')}\n"
        f"{r.get('body', '')}\n"
        f"URL: {r.get('href', '')}"
        for i, r in enumerate(achados, 1)
    )
    resultado += ("\n\n(Estes são apenas resumos. Para tabelas, agendas ou "
                  "listas completas, chame abrir_pagina em uma das URLs.)")
    _cache_busca[chave] = resultado
    return resultado


def abrir_pagina(url: str) -> str:
    """Baixa uma página web e retorna seu conteúdo em texto.

    Use para ler tabelas de jogos, calendários, agendas, listas e artigos
    completos que não aparecem nos resumos de buscar_web.

    Args:
        url: o endereço completo da página, começando com http ou https.

    Returns:
        O texto extraído da página, truncado em cerca de 8000 caracteres.
    """
    url = url.strip().strip('<>"\'')
    if not url.startswith(("http://", "https://")):
        return "URL inválida. Forneça um endereço iniciando com http ou https."

    if url in _cache_pagina:
        log.info("📄 (cache) %s", url[:70])
        return _cache_pagina[url]

    log.info("📄 abrindo: %s", url[:70])
    try:
        r = requests.get(url, headers=CABECALHO_HTTP, timeout=25)
        r.raise_for_status()
        if "html" not in r.headers.get("Content-Type", "") and not r.text:
            return "A página não retornou conteúdo em texto."
        r.encoding = r.encoding or "utf-8"
        bruto = r.text
    except Exception as e:
        log.warning("   falha ao abrir: %s", type(e).__name__)
        return (f"Não foi possível abrir a página ({type(e).__name__}). "
                "Tente outra URL dos resultados.")

    texto = re.sub(r"(?is)<(script|style|noscript|svg|head)\b.*?</\1>", " ", bruto)
    texto = re.sub(r"(?i)</(tr|p|div|li|h[1-6]|table)>", "\n", texto)
    texto = re.sub(r"(?i)</t[dh]>", " | ", texto)
    texto = re.sub(r"<[^>]+>", " ", texto)
    texto = html.unescape(texto)
    texto = re.sub(r"[ \t\xa0]+", " ", texto)
    texto = re.sub(r"\n\s*\n+", "\n", texto).strip()

    if len(texto) < 120:
        resultado = ("A página não expôs texto útil (provavelmente renderizada "
                     "por JavaScript). Tente outra fonte.")
    else:
        limite = 8000
        resultado = texto[:limite]
        if len(texto) > limite:
            resultado += f"\n\n[...truncado, {len(texto)} caracteres no total]"
        log.info("   %d caracteres extraídos", len(texto))

    _cache_pagina[url] = resultado
    return resultado


def cotacao_moeda(par: str) -> str:
    """Consulta a cotação atual de moedas e criptomoedas em tempo real.

    Args:
        par: par no formato ORIGEM-DESTINO, por exemplo 'USD-BRL' para dólar,
            'EUR-BRL' para euro, 'BTC-BRL' para bitcoin.

    Returns:
        Compra, venda, variação do dia, máxima, mínima e data da cotação.
    """
    par = par.upper().strip().replace("/", "-").replace("_", "-")
    if "-" not in par:
        par = f"{par}-BRL"

    log.info("💱 cotação: %s", par)
    try:
        r = requests.get(
            f"https://economia.awesomeapi.com.br/json/last/{par}", timeout=15
        )
        r.raise_for_status()
        d = next(iter(r.json().values()))
    except Exception as e:
        log.warning("   cotação falhou: %s", type(e).__name__)
        return f"Não foi possível obter a cotação de {par} ({type(e).__name__})."

    log.info("   %s = %s", par, d.get("bid"))
    return (
        f"{d.get('name', par)}\n"
        f"Compra: {d.get('bid')}\n"
        f"Venda: {d.get('ask')}\n"
        f"Variação no dia: {d.get('pctChange')}%\n"
        f"Máxima: {d.get('high')} | Mínima: {d.get('low')}\n"
        f"Atualizado em: {d.get('create_date')}\n"
        f"Fonte: AwesomeAPI (mercado em tempo real)"
    )


FERRAMENTAS = [hoje, buscar_web, abrir_pagina, cotacao_moeda]


@dataclass
class Config:
    api_key: str = os.getenv("GEMINI_API_KEY", "")
    modelo: str = ""
    usar_busca: bool = True
    intervalo_min: float = 15.0
    max_tentativas: int = 3
    backoff_base: float = 20.0
    temperatura: float = 0.4
    max_chamadas_ferramenta: int = 10
    candidatos: list = field(default_factory=lambda: list(CANDIDATOS))

    def __post_init__(self):
        if not self.api_key:
            raise RuntimeError("Defina GEMINI_API_KEY no arquivo .env")


# ------------------------------------------------------------------ UTILITÁRIOS
def classificar_429(erro: Exception) -> tuple[str, bool]:
    texto = str(erro)
    if "per_minute" in texto:
        return ("RPM — limite por minuto", True)
    if "input_token" in texto:
        return ("TPM — tokens por minuto", True)
    if "per_day" in texto:
        return ("RPD — limite diário do modelo", False)
    if "QuotaFailure" in texto or "retryDelay" in texto:
        return ("cota transitória", True)
    return ("COTA DE PROJETO — sem free tier nesta chave", False)


def calcular_pausa(erro: Exception, tentativa: int, base: float) -> float:
    texto = str(erro)
    if "retryDelay" in texto:
        try:
            trecho = texto.split("retryDelay")[1][:24]
            digitos = "".join(c for c in trecho if c.isdigit())
            if digitos:
                return float(digitos[:3]) + random.uniform(0, 2)
        except (IndexError, ValueError):
            pass
    return base * (2 ** (tentativa - 1)) + random.uniform(0, 3)


def dormir(segundos: float) -> bool:
    try:
        fim = time.monotonic() + segundos
        while True:
            resta = fim - time.monotonic()
            if resta <= 0:
                return True
            time.sleep(min(0.5, resta))
    except KeyboardInterrupt:
        log.warning("Espera cancelada pelo usuário.")
        return False


# ---------------------------------------------------------------------- AGENTE
class AgentePesquisa:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.cliente = genai.Client(api_key=cfg.api_key)
        self.busca_ativa = cfg.usar_busca
        self.ultima_chamada = 0.0
        self.modelo_atual = cfg.modelo or self._detectar_modelo()
        self.chat = self._novo_chat()

    def _detectar_modelo(self) -> str:
        log.info("Preflight: procurando modelo disponível...")
        for nome in self.cfg.candidatos:
            try:
                self.cliente.models.generate_content(
                    model=nome,
                    contents="ok",
                    config=types.GenerateContentConfig(
                        temperature=0, max_output_tokens=8
                    ),
                )
                log.info("✓ modelo %s OK", nome)
                return nome
            except genai_errors.ClientError as e:
                log.warning("✗ %s → erro %s", nome, getattr(e, "code", "?"))
            except Exception as e:
                log.warning("✗ %s → %s", nome, type(e).__name__)

        log.error("Nenhum modelo disponível. Verifique sua chave em "
                  "https://aistudio.google.com/apikey")
        sys.exit(1)

    def _novo_chat(self):
        return self.cliente.chats.create(
            model=self.modelo_atual,
            config=types.GenerateContentConfig(
                temperature=self.cfg.temperatura,
                tools=FERRAMENTAS if self.busca_ativa else None,
                automatic_function_calling=types.AutomaticFunctionCallingConfig(
                    maximum_remote_calls=self.cfg.max_chamadas_ferramenta
                ) if self.busca_ativa else None,
                system_instruction=(
                    instrucao_com_busca() if self.busca_ativa
                    else instrucao_sem_busca()
                ),
            ),
        )

    def _aguardar_janela(self) -> bool:
        espera = self.cfg.intervalo_min - (time.monotonic() - self.ultima_chamada)
        if espera > 0:
            log.info("Throttle: %.1fs para respeitar o RPM", espera)
            return dormir(espera)
        return True

    def perguntar(self, pergunta: str) -> str | None:
        for tentativa in range(1, self.cfg.max_tentativas + 1):
            if not self._aguardar_janela():
                return None
            try:
                resposta = self.chat.send_message(pergunta)
                self.ultima_chamada = time.monotonic()
                self._logar_uso(resposta)
                return (resposta.text or "").strip()

            except genai_errors.ClientError as e:
                self.ultima_chamada = time.monotonic()

                if getattr(e, "code", None) != 429:
                    log.error("Erro %s: %s", getattr(e, "code", "?"), e)
                    return None

                descricao, repetir = classificar_429(e)
                log.warning("429 (%d/%d) → %s", tentativa,
                            self.cfg.max_tentativas, descricao)

                if not repetir:
                    log.error("Cota não recuperável por retry. Abortando.")
                    log.error("Monitore em: https://ai.dev/rate-limit")
                    return None
                if tentativa == self.cfg.max_tentativas:
                    log.error("Tentativas esgotadas.")
                    return None
                if not dormir(calcular_pausa(e, tentativa, self.cfg.backoff_base)):
                    return None

            except genai_errors.ServerError as e:
                log.warning("Erro 5xx: %s", e)
                if not dormir(self.cfg.backoff_base * tentativa):
                    return None

            except KeyboardInterrupt:
                log.warning("Requisição interrompida.")
                return None

            except Exception as e:
                log.error("Falha inesperada [%s]: %s", type(e).__name__, e)
                return None

        return None

    @staticmethod
    def _logar_uso(resposta) -> None:
        uso = getattr(resposta, "usage_metadata", None)
        if uso:
            log.info("Tokens → in:%s out:%s total:%s",
                     uso.prompt_token_count,
                     uso.candidates_token_count,
                     uso.total_token_count)

    def resetar(self) -> None:
        _cache_busca.clear()
        _cache_pagina.clear()
        self.chat = self._novo_chat()
        log.info("Contexto e caches reiniciados.")


# ------------------------------------------------------------------------ CLI
def main() -> None:
    try:
        agente = AgentePesquisa(Config())
    except RuntimeError as e:
        print(f"Erro de configuração: {e}")
        return

    print("\n=== Agente de Pesquisa (Gemini + Web + Leitura + Câmbio) ===")
    print(f"Modelo: {agente.modelo_atual} | "
          f"Ferramentas: {'on' if agente.busca_ativa else 'off'}")
    print(f"Contexto temporal: {contexto_temporal()}")
    print("Comandos: /sair  /limpar  /busca  /modelo <nome>  /status\n")

    while True:
        try:
            entrada = input("Você → ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nEncerrando.")
            break

        if not entrada:
            continue
        if entrada == "/sair":
            print("Até logo, Luivan.")
            break
        if entrada == "/limpar":
            agente.resetar()
            continue
        if entrada == "/busca":
            agente.busca_ativa = not agente.busca_ativa
            agente.resetar()
            print(f"[ferramentas: {'ATIVAS' if agente.busca_ativa else 'DESATIVADAS'}]")
            continue
        if entrada.startswith("/modelo "):
            agente.modelo_atual = entrada.split(maxsplit=1)[1]
            agente.resetar()
            print(f"[modelo: {agente.modelo_atual}]")
            continue
        if entrada == "/status":
            print(f"[modelo={agente.modelo_atual} | "
                  f"ferramentas={agente.busca_ativa} | "
                  f"buscas em cache={len(_cache_busca)} | "
                  f"páginas em cache={len(_cache_pagina)}]")
            continue

        resposta = agente.perguntar(entrada)
        print(f"\nAgente → {resposta or '(sem resposta — veja o log acima)'}\n")


if __name__ == "__main__":
    main()
