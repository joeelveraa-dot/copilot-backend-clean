# backend/fae_core/llm.py
import os, json, logging
from openai import AzureOpenAI

log = logging.getLogger("fae_core.llm")

AZURE_KEY        = os.getenv("AZURE_OPENAI_KEY")
AZURE_ENDPOINT   = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")
AZURE_VERSION    = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01")

client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

def call_llm(messages: list[dict], **kwargs) -> str:
    """
    Wrapper simple para chat.completions de Azure OpenAI.
    """
    try:
        resp = client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=messages,
            **kwargs
        )
        return resp.choices[0].message.content
    except Exception as e:
        log.exception("LLM error: %s", e)
        raise

def extract_first_json_block(text: str) -> dict:
    """
    Extrae el primer bloque JSON balanceado {...} de un texto.
    Más robusto que un regex {.*}.
    """
    stack, start = 0, -1
    for i, ch in enumerate(text):
        if ch == '{':
            if stack == 0:
                start = i
            stack += 1
        elif ch == '}':
            stack -= 1
            if stack == 0 and start >= 0:
                return json.loads(text[start:i+1])
    raise ValueError("No se encontró JSON válido en la respuesta")
