from config import Config
from agent import AgentePesquisa, contexto_temporal
from tools.ferramentas import _cache_busca, _cache_pagina

def iniciar_cli() -> None:
    try:
        agente = AgentePesquisa(Config())
    except RuntimeError as e:
        print(f"Erro de configuração: {e}")
        return

    print("\n=== Agente de Pesquisa (Modularizado) ===")
    print(f"Modelo: {agente.modelo_atual} | Ferramentas: {'on' if agente.busca_ativa else 'off'}")
    print(f"Contexto temporal: {contexto_temporal()}")
    print("Comandos: /sair  /limpar  /busca  /modelo <nome>  /status\n")

    while True:
        try:
            entrada = input("Você → ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nEncerrando.")
            break

        if not entrada: continue
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
            print(f"[modelo={agente.modelo_atual} | ferramentas={agente.busca_ativa} | "
                  f"buscas em cache={len(_cache_busca)} | páginas em cache={len(_cache_pagina)}]")
            continue

        resposta = agente.perguntar(entrada)
        print(f"\nAgente → {resposta or '(sem resposta — veja o log acima)'}\n")
