import os
from dotenv import load_dotenv

load_dotenv()

KEY_BY_PROVIDER = {
    "gemini": "GEMINI_API_KEY",
    "xai": "XAI_API_KEY",
    "openai": "OPENAI_API_KEY",
}

def main() -> None:
    provider = os.getenv("LLM_PROVIDER", "gemini").strip().lower()
    var_name = KEY_BY_PROVIDER.get(provider)

    if not var_name:
        print(f"ERRO: provider desconhecido: {provider}")
        return

    api_key = os.getenv(var_name)
    if not api_key:
        print(f"ERRO: {var_name} não encontrada no .env")
        return

    print(f"Ambiente OK | provider: {provider} | chave: ...{api_key[-4:]}")

if __name__ == "__main__":
    main()
