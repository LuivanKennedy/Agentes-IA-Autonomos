import sys
import time
import random
import os
import re
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
        "Você é um agente autônomo de inteligência, criação e pesquisa. Responda em português do Brasil.\n\n"
        f"DATA E HORA ATUAIS: {contexto_temporal()}.\n\n"
        "FERRAMENTAS DE CRIAÇÃO E AÇÃO:\n"
        "- criar_documento_word: Use para gerar documentos. REGRA OBRIGATÓRIA: Use formatação Markdown (inicie com '#' para Títulos, '##' para Subtítulos e use '**' para palavras em negrito). Para incluir imagens geradas no documento, você DEVE escrever no texto a tag [IMAGEM: nome_do_arquivo.jpg] isolada em uma linha, no local exato onde deseja a imagem.\n"
        "- criar_pdf: Use para gerar relatórios definitivos e limpos.\n"
        "- criar_planilha_excel: Use se o usuário pedir tabelas, planilhas ou organizar dados. Converta os dados em JSON e envie para a ferramenta.\n"
        "- gerar_imagem: Use se o usuário pedir para gerar, criar ou desenhar uma imagem avulsa. Invente um prompt rico em inglês para a ferramenta.\n"
        "- criar_apresentacao_com_ia: OBRIGATÓRIO quando o usuário pedir PPT/apresentações. Você deve planejar o roteiro, inventar prompts de imagem (em inglês) para cada slide, formatar um JSON estrito (sem erros de aspas) e chamar a ferramenta.\n\n"
        "FERRAMENTAS DE PESQUISA (ESTRATÉGIA PROGRESSIVA):\n"
        "- ler_arquivo: Extrai dados de documentos locais do usuário (.txt, .pdf, .xlsx, .pptx, imagens).\n"
        "- cotacao_moeda: use SEMPRE para dólar, euro, bitcoin.\n"
        "- buscar_web: OBRIGATÓRIA para iniciar pesquisas de internet. Se os resumos forem suficientes, RESPONDA.\n"
        "- abrir_pagina: Use para aprofundar se a busca for superficial.\n\n"
        "REGRAS:\n"
        "1. EVITE LOOPS: Faça no máximo 4 tentativas de ferramentas seguidas.\n"
        "2. Se criar um arquivo, avise o usuário onde está e comemore.\n"
        "3. HONESTIDADE: Nunca invente fatos em pesquisas. Mas seja altamente criativo ao gerar documentos.\n"
        "4. TOM E ESTRUTURA (ANTI-ROBÔ): Se o usuário pedir um documento 'oficial', 'executivo' ou 'profissional', escreva focando em estrutura de excelência (o sistema já cuidará da ABNT), porém SEMPRE mantenha uma linguagem natural e empática, fugindo do tom robótico. OBRIGATÓRIO: Evite o uso de travessões longos ('—') para intercalar explicações no meio das frases. Use vírgulas ou parênteses, que soam muito mais naturais para humanos. Nunca crie pontuações bizarras como travessão seguido de vírgula ('—,')."
    )

def instrucao_sem_busca() -> str:
    return (
        "Você é um agente autônomo. Responda em português. "
        f"Data atual: {contexto_temporal()}. "
        "Você NÃO tem acesso à internet nesta sessão. Não invente dados de pesquisa, mas pode criar documentos livremente."
    )

def classificar_erro_api(erro: Exception) -> tuple[str, bool]:
    texto = str(erro)
    if "per_minute" in texto: return ("Limites por Minuto", True)
    if "input_token" in texto: return ("Limite de Tokens", True)
    if "per_day" in texto: return ("Limite Diário Excedido", False)
    if "QuotaFailure" in texto or "retryDelay" in texto: return ("Cota Transitória", True)
    if "503" in texto or "high demand" in texto.lower() or "unavailable" in texto.lower():
        return ("Servidor Congestionado (503)", True)
    return ("Erro Desconhecido / Cota de Projeto", False)

def calcular_pausa(erro: Exception, tentativa: int, base: float) -> float:
    texto = str(erro)
    if "retryDelay" in texto:
        try:
            trecho = texto.split("retryDelay")[1][:24]
            digitos = "".join(c for c in trecho if c.isdigit())
            if digitos: return float(digitos[:3]) + random.uniform(0, 2)
        except (IndexError, ValueError): pass
    return base * (2 ** (tentativa - 1)) + random.uniform(2, 5)

def dormir(segundos: float) -> bool:
    try:
        fim = time.monotonic() + segundos
        while True:
            resta = fim - time.monotonic()
            if resta <= 0: return True
            time.sleep(min(0.5, resta))
    except KeyboardInterrupt:
        log.warning("Espera cancelada pelo usuário.")
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
        ferramentas_lista = []
        if self.busca_ativa and FERRAMENTAS:
            for nome_ferramenta, func in FERRAMENTAS.items():
                if callable(func):
                    try:
                        func.__name__ = nome_ferramenta
                    except Exception as e:
                        log.debug("Não foi possível alterar __name__ da ferramenta %s: %s", nome_ferramenta, e)
                    ferramentas_lista.append(func)
        
        try:
            config_auto = types.AutomaticFunctionCallingConfig(disable=True)
        except AttributeError:
            config_auto = {"disable": True}
            
        return self.cliente.chats.create(
            model=self.modelo_atual,
            config=types.GenerateContentConfig(
                temperature=self.cfg.temperatura,
                tools=ferramentas_lista if ferramentas_lista else None,
                automatic_function_calling=config_auto,
                system_instruction=instrucao_com_busca() if self.busca_ativa else instrucao_sem_busca(),
            ),
        )

    def perguntar(self, pergunta: str) -> str | None:
        padrao_midia = re.compile(r'([a-zA-Z0-9_.\-\\/:]+\.(?:jpg|jpeg|png|webp|heic|mp3|wav|ogg|mp4|avi|mov))', re.IGNORECASE)
        possiveis_arquivos = list(set(padrao_midia.findall(pergunta)))
        
        conteudos_para_enviar = [pergunta]
        
        for caminho in possiveis_arquivos:
            if os.path.exists(caminho):
                log.info("👁️ Preparando mídia (upload) para o agente: %s", caminho)
                try:
                    arquivo_upado = self.cliente.files.upload(file=caminho)
                    conteudos_para_enviar.append(arquivo_upado)
                except Exception as e:
                    log.warning("Falha ao preparar mídia %s: %s", caminho, type(e).__name__)

        for tentativa in range(1, self.cfg.max_tentativas + 1):
            espera = self.cfg.intervalo_min - (time.monotonic() - self.ultima_chamada)
            if espera > 0 and not dormir(espera): return None

            try:
                resposta = self.chat.send_message(conteudos_para_enviar)
                
                ciclos = 0
                while getattr(resposta, "function_calls", None) and ciclos < 10:
                    partes_resposta = []
                    
                    for fc in resposta.function_calls:
                        nome_ferramenta = fc.name
                        argumentos = fc.args or {}
                        log.info("⚙️  Agente acionou ferramenta: %s", nome_ferramenta)
                        
                        if nome_ferramenta in FERRAMENTAS:
                            func = FERRAMENTAS[nome_ferramenta]
                            if callable(func):
                                try:
                                    resultado = func(**argumentos)
                                    if not isinstance(resultado, dict):
                                        resultado = {"resultado": str(resultado)}
                                except Exception as err_func:
                                    log.error("Erro interno na ferramenta %s: %s", nome_ferramenta, err_func)
                                    resultado = {"erro": str(err_func)}
                            else:
                                erro_msg = f"A ferramenta '{nome_ferramenta}' não é uma função válida."
                                log.error(erro_msg)
                                resultado = {"erro": erro_msg}
                        else:
                            erro_msg = f"Ferramenta '{nome_ferramenta}' desconhecida."
                            log.error(erro_msg)
                            resultado = {"erro": erro_msg}
                            
                        partes_resposta.append(
                            types.Part.from_function_response(
                                name=nome_ferramenta,
                                response=resultado
                            )
                        )
                    
                    resposta = self.chat.send_message(partes_resposta)
                    ciclos += 1

                self.ultima_chamada = time.monotonic()
                uso = getattr(resposta, "usage_metadata", None)
                if uso:
                    log.info("Tokens → in:%s out:%s total:%s",
                             uso.prompt_token_count, uso.candidates_token_count, uso.total_token_count)
                
                texto_final = (resposta.text or "").strip()
                
                if not texto_final and getattr(resposta, "function_calls", None):
                    return "⚠️ O agente fez várias ações, mas não conseguiu gerar um texto final. Seja mais específico!"
                
                return texto_final

            except genai_errors.ClientError as e:
                self.ultima_chamada = time.monotonic()
                codigo_erro = getattr(e, "code", None)
                
                if codigo_erro not in (429, 503): 
                    log.error("Falha na API: %s", e)
                    return None
                
                descricao, repetir = classificar_erro_api(e)
                log.warning("Erro API %s (%d/%d) → %s. Tentando novamente em breve...", codigo_erro, tentativa, self.cfg.max_tentativas, descricao)
                
                if not repetir or tentativa == self.cfg.max_tentativas: 
                    return "A API do Google está indisponível no momento devido à alta demanda. Por favor, aguarde alguns minutos."
                
                if not dormir(calcular_pausa(e, tentativa, self.cfg.backoff_base)): return None
                
            except Exception as e:
                log.error("Falha inesperada: %s", repr(e)) 
                return None
                
        return None

    def resetar(self) -> None:
        _cache_busca.clear()
        _cache_pagina.clear()
        salvar_cache()
        self.chat = self._novo_chat()
        log.info("Contexto e caches reiniciados e salvos.")
