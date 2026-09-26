# -*- coding: utf-8 -*-
"""Cliente único para o AMALIA (LLM português) — dois back-ends, uma função.

O AMALIA não tem uma API pública "chave-na-mão"; há duas formas de lhe chamar,
ambas simples POST HTTP (ver https://amalia-llm.github.io/intro.html):

  1. IAedu  — API alojada em https://iaedu.pt para docentes/investigadores.
              Enviar {channel_id, thread_id, user_info, message} com o cabeçalho
              'x-api-key'. A resposta vem em linhas JSON (NDJSON); usa-se a do
              tipo "message".
  2. vLLM   — o próprio AMALIA servido localmente com vLLM, no formato da OpenAI
   (openai)   (/v1/chat/completions, cabeçalho 'Authorization: Bearer <chave>').
              Requer GPU (mín. NVIDIA A100 40GB). temperature=0 → determinista.

Como ainda não sabemos qual dos dois a EPAL vai usar, este módulo suporta os
dois. Configura-se por variáveis de ambiente e NÃO precisa de credenciais para
o resto do pipeline correr: se não estiver configurado, `is_configured()` devolve
False e o comparador escreve o Excel na mesma, deixando as colunas do AMALIA a
"PENDENTE" e guardando o prompt para se poder correr mais tarde (ou colar à mão).

Variáveis de ambiente
---------------------
  AMALIA_BACKEND   'iaedu' | 'openai' | 'vllm'   (vazio = modo dry-run)
  AMALIA_ENDPOINT  URL do endpoint
  AMALIA_API_KEY   chave / token
  AMALIA_MODEL     nome do modelo (openai/vllm)  [def: amalia-llm/AMALIA-9B-0626-DPO]
  AMALIA_CHANNEL_ID / AMALIA_THREAD_ID           (opcional, para o IAedu)
  AMALIA_TIMEOUT   segundos                        [def: 900]
"""
from __future__ import annotations
import json
import os
import uuid

import ssl

try:
    import requests
    from requests.adapters import HTTPAdapter
except ImportError:                                # requests só é preciso quando há credenciais
    requests = None
    HTTPAdapter = object


DEFAULT_MODEL = "amalia-llm/AMALIA-9B-0626-DPO"


# --------------------------------------------------------------------------- #
#  TLS: a rede da EPAL passa por um proxy cujo certificado de CA tem o campo
#  "Basic Constraints" NÃO marcado como crítico. Os browsers aceitam-no; o
#  Python 3.13 activa VERIFY_X509_STRICT por omissão e rejeita-o com
#      CERTIFICATE_VERIFY_FAILED: Basic Constraints of CA cert not marked critical
#  — um erro que não tem nada a ver com a chave da API, mas que aparece como se
#  tivesse. Relaxamos ESSA regra e mais nenhuma:
#      * a cadeia continua a ser verificada contra o arquivo de certificados;
#      * o nome do servidor continua a ser verificado.
#  NÃO é o mesmo que verify=False, que enviaria a chave por uma ligação que
#  deixámos de autenticar. Para voltar ao comportamento estrito: AMALIA_TLS_STRICT=1
# --------------------------------------------------------------------------- #
def _tls_context() -> "ssl.SSLContext":
    ctx = ssl.create_default_context()
    if (os.getenv("AMALIA_TLS_STRICT") or "").strip() not in {"1", "true", "sim"}:
        ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return ctx


class _TlsAdapter(HTTPAdapter):
    def init_poolmanager(self, *a, **kw):
        kw["ssl_context"] = _tls_context()
        return super().init_poolmanager(*a, **kw)


def _session():
    s = requests.Session()
    s.mount("https://", _TlsAdapter())
    return s


class AmaliaNotConfigured(RuntimeError):
    """Levantada quando se tenta chamar o AMALIA sem back-end configurado."""


def _cfg():
    return {
        "backend": (os.getenv("AMALIA_BACKEND") or "").strip().lower(),
        "endpoint": (os.getenv("AMALIA_ENDPOINT") or "").strip(),
        "api_key": (os.getenv("AMALIA_API_KEY") or "").strip(),
        "model": (os.getenv("AMALIA_MODEL") or DEFAULT_MODEL).strip(),
        "channel_id": (os.getenv("AMALIA_CHANNEL_ID") or "").strip(),
        "thread_id": (os.getenv("AMALIA_THREAD_ID") or "").strip(),
        "timeout": float(os.getenv("AMALIA_TIMEOUT") or "900"),
        # 900 s, nao 120: um modelo local em CPU gera ~1 token/s, e uma
        # comparacao de criterios sao varias centenas de tokens.
    }


def is_configured() -> bool:
    """True quando há back-end + endpoint + chave para chamar o modelo."""
    c = _cfg()
    return bool(c["backend"] in {"iaedu", "openai", "vllm"} and c["endpoint"] and c["api_key"])


def backend_name() -> str:
    return _cfg()["backend"] or "(não configurado)"


# --------------------------------------------------------------------------- #
#  Back-end 1: IAedu (API alojada)
# --------------------------------------------------------------------------- #
def _ask_iaedu(message: str, system: str | None, c: dict) -> str:
    if requests is None:
        raise AmaliaNotConfigured("O pacote 'requests' não está instalado.")
    # O IAedu não tem papel 'system' explícito no exemplo da doc — antepomos as
    # instruções à mensagem do utilizador.
    full = (system + "\n\n" + message) if system else message
    form = {
        "channel_id": c["channel_id"] or str(uuid.uuid4()),
        "thread_id": c["thread_id"] or str(uuid.uuid4()),
        "user_info": "{}",
        "message": full,
    }
    r = _session().post(c["endpoint"], headers={"x-api-key": c["api_key"]},
                        data=form, timeout=c["timeout"])
    r.raise_for_status()
    # Resposta em linhas JSON separadas por linha em branco; usar a 1.ª "message".
    for chunk in r.text.split("\n\n"):
        chunk = chunk.strip()
        if not chunk:
            continue
        try:
            j = json.loads(chunk)
        except json.JSONDecodeError:
            continue
        if j.get("type") == "message":
            content = j.get("content", {})
            return content.get("content", "") if isinstance(content, dict) else str(content)
    return r.text.strip()


# --------------------------------------------------------------------------- #
#  Back-end 2: vLLM / formato OpenAI (auto-alojado)
# --------------------------------------------------------------------------- #
def _ask_openai(message: str, system: str | None, c: dict,
                temperature: float, max_tokens: int) -> str:
    if requests is None:
        raise AmaliaNotConfigured("O pacote 'requests' não está instalado.")
    url = c["endpoint"].rstrip("/")
    if url.endswith("/chat/completions"):
        pass                                       # já é o URL completo
    elif url.endswith("/v1"):
        url = url + "/chat/completions"            # ex.: http://localhost:11434/v1  (Ollama)
    else:
        url = url + "/v1/chat/completions"         # ex.: http://127.0.0.1:8001      (vLLM)
    messages = ([{"role": "system", "content": system}] if system else []) + \
               [{"role": "user", "content": message}]
    payload = {
        "model": c["model"],
        "messages": messages,
        "temperature": temperature,
        # "max_tokens", nao "max_completion_tokens": o vLLM e o TGI (HuggingFace)
        # so conhecem o nome antigo. Com o nome novo, o limite era ignorado.
        "max_tokens": max_tokens,
    }
    headers = {"Content-Type": "application/json"}
    if c.get("api_key"):
        headers["Authorization"] = f"Bearer {c['api_key']}"
                                      "Authorization": f"Bearer {c['api_key']}"},
                        json=payload, timeout=c["timeout"])
    r.raise_for_status()
    escolha = r.json()["choices"][0]
    # Uma resposta cortada a meio ainda tem os primeiros campos bem formados, e o
    # parser do comparador aceitava-a como completa. Melhor dize-lo em voz alta.
    if escolha.get("finish_reason") == "length":
        raise RuntimeError(
            "resposta truncada pelo limite de %d tokens (max_tokens) - "
            "aumente AMALIA_MAX_TOKENS" % max_tokens)
    return escolha["message"]["content"]


# --------------------------------------------------------------------------- #
#  Ponto de entrada único
# --------------------------------------------------------------------------- #
def ask(message: str, system: str | None = None,
        temperature: float = 0.0, max_tokens: int = 320) -> str:
    # ponytail: 320 e' o tecto que CABE no timeout. A 0,8 tokens/s (CPU, sem
    # GPU) 900 tokens sao 1125 s -- mais do que os 900 s de timeout, por isso a
    # chamada morria SEMPRE antes de o modelo acabar. Com GPU pode subir-se via
    # AMALIA_MAX_TOKENS.
    """Envia um prompt ao AMALIA e devolve o texto da resposta.

    Levanta AmaliaNotConfigured se não houver back-end configurado — o chamador
    deve verificar is_configured() primeiro se quiser um modo dry-run silencioso.
    """
    c = _cfg()
    if not is_configured():
        raise AmaliaNotConfigured(
            "AMALIA não configurado. Defina AMALIA_BACKEND (iaedu|openai), "
            "AMALIA_ENDPOINT e AMALIA_API_KEY.")
    if c["backend"] == "iaedu":
        return _ask_iaedu(message, system, c)
    return _ask_openai(message, system, c, temperature, max_tokens)


if __name__ == "__main__":                         # teste rápido de fumo
    print("Back-end:", backend_name(), "| configurado:", is_configured())
    if is_configured():
        print(ask("Responde apenas com: OK", max_tokens=10))
