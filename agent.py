import sys
import time
import random
from datetime import datetime
from google import genai
from google.genai import types
from google.genai import errors as genai_errors

from config import log, Config, FUSO, DIAS
from tools.ferramentas import FERRAMENTAS, salvar_cache, _cache_busca, _cache_pagina

def contexto_temporal() -> str:
    agora = datetime.now(FUSO)
    return f"{DIAS[agora.weekday()]}, {agora:%d/%m/%Y} às {agora:%H:%M} (horário de Brasília)"

def instrucao_com_busca() -> str:
    return (
        "Você é um agente de pesquisa objetivo. Responda em português do Brasil.\n\n"
        f"DATA E HORA ATUAIS: {contexto_temporal()}.\n\n"
        "FERRAMENTAS (ESTRATÉGIA PROGRESSIVA):\n"
        "- hoje: confirme a data antes de raciocínio temporal.\n"
        "- cotacao_moeda: use SEMPRE para dólar, euro, bitcoin. Não use buscar_web para isso.\n"
        "- buscar_web: OBRIGATÓRIA para iniciar pesquisas. Ela retorna RESUMOS rápidos. "
        "Se a pergunta for SIMPLES (ex: idade, capital, placar final) e a resposta estiver "
        "nos resumos, RESPONDA IMEDIATAMENTE. Não use outras ferramentas.\n"
        "- abrir_pagina: use APENAS se a pergunta for COMPLEXA (ex: calendários, tabelas completas, "
        "escalações detalhadas) ou se os resumos do buscar_web forem insuficientes.\n\n"
        "REGRAS DE COMPORTAMENTO (MUITO IMPORTANTE):\n"
        "1. EVITE LOOPS: Faça no MÁXIMO 4 ou 5 tentativas de busca ou abrir página por pergunta.\n"
        "2. Se não encontrar a resposta exata após algumas tentativas, PARE IMEDIATAMENTE de usar ferramentas.\n"
        "3. Responda com as informações que conseguiu reunir até o momento ou admita claramente que a informação não está disponível/acessível.\n\n"
        "HONESTIDADE: Cite os links usados. Nunca invente dados."
    )

def instrucao_sem_busca() -> str:
    return (
        "Você é um agente de pesquisa. Responda em português. "
        f"Data atual: {contexto_temporal()}. "
        "Você NÃO tem acesso à internet nesta sessão. Não invente dados."
    )

def classificar_429(erro: Exception) -> tuple[str, bool]:
    texto = str(erro)
    if "per_minute" in texto: return ("RPM", True)
    if "input_token" in texto: return ("TPM", True)
    if "per_day" in texto: return ("RPD", False)
    if "QuotaFailure" in texto or "retryDelay" in texto: return ("cota transitória", True)
    return ("COTA DE PROJETO", False)

def calcular_pausa(erro: Exception, tentativa: int, base: float) -> float:
    texto = str(erro)
    if "retryDelay" in texto:
        try:
            trecho = texto.split("retryDelay")[1][:24]
            digitos = "".join(c for c in trecho if c.isdigit())
            if digitos: return float(digitos[:3]) + random.uniform(0, 2)
        except (IndexError, ValueError): pass
    return base * (2 ** (tentativa - 1)) + random.uniform(0, 3)

def dormir(segundos: float) -> bool:
    try:
        fim = time.monotonic() + segundos
        while True:
            resta = fim - time.monotonic()
            if resta <= 0: return True
            time.sleep(min(0.5, resta))
    except KeyboardInterrupt:
        log.warning("Espera cancelada.")
        return False

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
                    model=nome, contents="ok",
                    config=types.GenerateContentConfig(temperature=0, max_output_tokens=8)
                )
                log.info("✓ modelo %s OK", nome)
                return nome
            except Exception as e:
                log.warning("✗ %s falhou", nome)
        log.error("Nenhum modelo disponível.")
        sys.exit(1)

    def _novo_chat(self):
        return self.cliente.chats.create(
            model=self.modelo_atual,
            config=types.GenerateContentConfig(
                temperature=self.cfg.temperatura,
                tools=FERRAMENTAS if self.busca_ativa else None,
                automatic_function_calling=types.AutomaticFunctionCallingConfig(
                    disable=False,
                    maximum_remote_calls=10  # Aumentamos o limite de ferramentas por turno!
                ) if self.busca_ativa else None,
                system_instruction=instrucao_com_busca() if self.busca_ativa else instrucao_sem_busca(),
            ),
        )

    def perguntar(self, pergunta: str) -> str | None:
        for tentativa in range(1, self.cfg.max_tentativas + 1):
            espera = self.cfg.intervalo_min - (time.monotonic() - self.ultima_chamada)
            if espera > 0 and not dormir(espera): return None

            try:
                resposta = self.chat.send_message(pergunta)
                self.ultima_chamada = time.monotonic()
                uso = getattr(resposta, "usage_metadata", None)
                if uso:
                    log.info("Tokens → in:%s out:%s total:%s",
                             uso.prompt_token_count, uso.candidates_token_count, uso.total_token_count)
                
                # Pegamos o texto. Se vier vazio, garantimos que não deu erro por causa das ferramentas.
                texto_final = (resposta.text or "").strip()
                
                # Rede de segurança: avisa o usuário caso o agente atinja o limite antes de responder
                if not texto_final and getattr(resposta, "function_calls", None):
                    return "⚠️ Fiz várias pesquisas, mas atingi o limite de ações automáticas antes de concluir a resposta. Tente perguntar de forma mais específica!"
                
                return texto_final

            except genai_errors.ClientError as e:
                self.ultima_chamada = time.monotonic()
                if getattr(e, "code", None) != 429: return None
                descricao, repetir = classificar_429(e)
                log.warning("429 (%d/%d) → %s", tentativa, self.cfg.max_tentativas, descricao)
                if not repetir or tentativa == self.cfg.max_tentativas: return None
                if not dormir(calcular_pausa(e, tentativa, self.cfg.backoff_base)): return None
            except Exception as e:
                log.error("Falha inesperada: %s", e)
                return None
        return None

    def resetar(self) -> None:
        _cache_busca.clear()
        _cache_pagina.clear()
        salvar_cache()
        self.chat = self._novo_chat()
        log.info("Contexto e caches reiniciados e salvos.")
