# -*- coding: utf-8 -*-
r"""Configuracao central do projeto EPAL - resolve caminhos e definicoes por maquina.

PORQUE EXISTE ESTE FICHEIRO
---------------------------
A pasta do projeto e partilhada (SharePoint / Teams). Cada pessoa sincroniza-a
para um caminho DIFERENTE no seu PC:

    C:\Users\ana\EPAL\EPAL-project
    C:\Users\rui\OneDrive - ...\Bureau\EPAL\EPAL-project

Logo NADA dentro do projeto pode ter caminhos absolutos escritos a mao. Duas regras:

  1. Todos os caminhos DO PROJETO sao derivados deste ficheiro (ROOT). Assim
     acompanham a pasta para onde quer que ela seja sincronizada.

  2. As definicoes PROPRIAS DE CADA MAQUINA (endpoint do AMALIA, chaves, portas)
     NAO podem viver dentro da pasta partilhada - seriam sincronizadas para toda
     a gente e cada PC sobrepunha-se ao anterior. Vivem em:

         %LOCALAPPDATA%\EPAL\epal.local.ini

     que e local a cada PC e nunca e sincronizado.

Ordem de precedencia (a primeira que existir ganha):

     variavel de ambiente  >  epal.local.ini  >  valor por omissao

DIAGNOSTICO
-----------
Se algo nao funcionar num PC, peca a essa pessoa para correr:

     python src\epal_config.py

e enviar o texto que aparece. Mostra tudo o que este PC resolveu.
"""
from __future__ import annotations

import configparser
import os

# --------------------------------------------------------------------------- #
#  1. Caminhos do projeto - SEMPRE derivados da localizacao deste ficheiro.
#     Nunca escrever caminhos absolutos aqui.
# --------------------------------------------------------------------------- #
HERE = os.path.dirname(os.path.abspath(__file__))          # .../EPAL-project/src
ROOT = os.path.dirname(HERE)                               # .../EPAL-project

DATA          = os.path.join(ROOT, "data")
DATA_POWERBI  = os.path.join(DATA, "powerbi")
DROP          = os.path.join(ROOT, "Drop_New_Licenses")
DROP_TUA      = os.path.join(DROP, "new_TUA_licences")
DROP_LURH     = os.path.join(DROP, "new_LURH_licences")
REPORTS       = os.path.join(ROOT, "reports")
OUTPUTS       = os.path.join(ROOT, "outputs")

DASHBOARD_DIR = os.path.join(REPORTS, "EPAL_Licencas_Dashboard")
DASHBOARD_TMDL = os.path.join(
    DASHBOARD_DIR, "EPAL_Licencas_Dashboard.SemanticModel",
    "definition", "expressions.tmdl")


# --------------------------------------------------------------------------- #
#  2. Definicoes por maquina - fora da pasta partilhada.
# --------------------------------------------------------------------------- #
def local_dir() -> str:
    """Pasta local a este PC (nunca sincronizada)."""
    base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "EPAL")


def local_config_path() -> str:
    return os.path.join(local_dir(), "epal.local.ini")


def local_venv_python() -> str:
    """O python do ambiente virtual desta maquina (criado pelo Automation files/CONFIGURAR_ESTE_PC.bat)."""
    return os.path.join(local_dir(), "venv", "Scripts", "python.exe")


_CACHE: configparser.ConfigParser | None = None


def _ini() -> configparser.ConfigParser:
    global _CACHE
    if _CACHE is None:
        cp = configparser.ConfigParser()
        path = local_config_path()
        if os.path.isfile(path):
            try:
                cp.read(path, encoding="utf-8")
            except Exception:
                pass                                    # ficheiro invalido -> usar omissoes
        _CACHE = cp
    return _CACHE


def get(key: str, default: str = "", section: str = "amalia",
        env: str | None = None) -> str:
    """Le uma definicao: variavel de ambiente > epal.local.ini > omissao."""
    if env:
        v = os.getenv(env)
        if v and v.strip():
            return v.strip()
    try:
        v = _ini().get(section, key)
        if v and v.strip():
            return v.strip()
    except Exception:
        pass
    return default


# --------------------------------------------------------------------------- #
#  3. AMALIA - as definicoes que mudam de maquina para maquina.
#     amalia_client.py le variaveis de ambiente; aqui preenchemo-las a partir
#     do ficheiro local, sem alterar o cliente.
# --------------------------------------------------------------------------- #
AMALIA_SETTINGS = {                     # chave no .ini -> variavel de ambiente
    "backend":    "AMALIA_BACKEND",
    "endpoint":   "AMALIA_ENDPOINT",
    "api_key":    "AMALIA_API_KEY",
    "model":      "AMALIA_MODEL",
    "channel_id": "AMALIA_CHANNEL_ID",
    "thread_id":  "AMALIA_THREAD_ID",
    "timeout":    "AMALIA_TIMEOUT",
}


def apply_amalia_env() -> None:
    """Copia as definicoes do AMALIA do .ini local para o ambiente do processo.

    Uma variavel de ambiente ja definida NUNCA e sobreposta - permite testar
    outro endpoint numa consola sem mexer no ficheiro.
    """
    for key, env_var in AMALIA_SETTINGS.items():
        if os.getenv(env_var):
            continue
        value = get(key, section="amalia")
        if value:
            os.environ[env_var] = value


TEMPLATE = """\
; ===================================================================
;  Definicoes locais do EPAL - SO DESTE PC.
;  Este ficheiro NAO esta na pasta partilhada, por isso pode ser
;  diferente em cada computador sem afetar os colegas.
;  Para o projeto so extrair licencas, nao e preciso mexer em nada.
; ===================================================================

[amalia]
; Preencher apenas quando o AMALIA (LLM) estiver instalado neste PC.
; Deixar em branco = a comparacao de criterios corre na mesma e escreve
; "PENDENTE" nas colunas do AMALIA.
;
; vLLM local:   backend = openai   e   endpoint = http://localhost:8000/v1
; IAedu:        backend = iaedu    e   endpoint = <url do IAedu>
backend  =
endpoint =
api_key  =
model    = amalia-llm/AMALIA-9B-0626-DPO
timeout  = 900
"""


def ensure_local_config() -> str:
    """Cria epal.local.ini com o modelo por omissao, se ainda nao existir."""
    path = local_config_path()
    if not os.path.isfile(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(TEMPLATE)
    return path


# --------------------------------------------------------------------------- #
#  4. Diagnostico
# --------------------------------------------------------------------------- #
def _status(path: str) -> str:
    return "OK    " if os.path.exists(path) else "EM FALTA"


def main() -> int:
    import sys
    print("=" * 70)
    print(" EPAL - diagnostico de configuracao desta maquina")
    print("=" * 70)
    print(f"  PC                  : {os.getenv('COMPUTERNAME', '?')}")
    print(f"  Python              : {sys.version.split()[0]}  ({sys.executable})")
    print()
    print("  PASTAS DO PROJETO (derivadas, nunca fixas)")
    for label, path in [("ROOT", ROOT), ("data/powerbi", DATA_POWERBI),
                        ("Drop TUA", DROP_TUA), ("Drop LURH", DROP_LURH),
                        ("reports", REPORTS), ("outputs", OUTPUTS)]:
        print(f"    [{_status(path)}] {label:<14} {path}")
    print()
    print("  DEFINICOES LOCAIS (fora da pasta partilhada)")
    cfg = local_config_path()
    print(f"    [{_status(cfg)}] epal.local.ini {cfg}")
    venv = local_venv_python()
    print(f"    [{_status(venv)}] venv local     {venv}")
    print()
    print("  AMALIA")
    apply_amalia_env()
    backend = os.getenv("AMALIA_BACKEND", "")
    if backend:
        key = os.getenv("AMALIA_API_KEY", "")
        print(f"    backend  : {backend}")
        print(f"    endpoint : {os.getenv('AMALIA_ENDPOINT', '')}")
        print(f"    api_key  : {'(definida)' if key else '(em falta)'}")
    else:
        print("    nao configurado - a comparacao de criterios escreve 'PENDENTE'.")
    print()
    print("  POWER BI")
    print(f"    [{_status(DASHBOARD_TMDL)}] expressions.tmdl")
    if os.path.isfile(DASHBOARD_TMDL):
        import re
        text = open(DASHBOARD_TMDL, encoding="utf-8").read()
        m = re.search(r'expression\s+DataFolder\s*=\s*"([^"]*)"', text)
        actual = m.group(1) if m else "(nao encontrado)"
        # O DataFolder normal e a ligacao estavel C:\EPAL\powerbi, igual em
        # todos os PCs. Comparar pelo caminho REAL: o que interessa e onde a
        # ligacao vai dar, nao como se escreve.
        try:
            ok = (os.path.normcase(os.path.realpath(actual))
                  == os.path.normcase(os.path.realpath(DATA_POWERBI)))
        except OSError:
            ok = False
        print(f"    DataFolder aponta para : {actual}")
        if ok and os.path.normcase(actual) != os.path.normcase(DATA_POWERBI):
            print(f"    ligacao estavel para   : {DATA_POWERBI}")
        elif not ok:
            print(f"    deveria dar em         : {DATA_POWERBI}")
        print(f"    -> {'CORRETO' if ok else 'ERRADO - corra Automation files/ATUALIZAR_POWERBI.bat'}")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
