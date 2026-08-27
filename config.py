import os
import logging
from datetime import timedelta, timezone
from dataclasses import dataclass, field
from dotenv import load_dotenv

load_dotenv()

# Configuração de Logs
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

# Configuração de Fuso Horário
try:
    from zoneinfo import ZoneInfo
    FUSO = ZoneInfo("America/Sao_Paulo")
except Exception:
    FUSO = timezone(timedelta(hours=-3), name="BRT")
    log.warning("tzdata ausente — usando UTC-3 fixo.")

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
