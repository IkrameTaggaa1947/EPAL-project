# -*- coding: utf-8 -*-
"""Confirma que TODAS as bibliotecas que o projeto usa importam mesmo.

PORQUE EXISTE
-------------
`python --version` responde na mesma com a pasta a meio da sincronizacao: o
python.exe ja ca esta, e as bibliotecas ainda nao. Quem configura o PC ve
"Python 3.13.1", fica descansado, e so descobre o problema quando a extracao
rebenta a meio com um ModuleNotFoundError.

O que parte a extracao e faltar UMA biblioteca. A unica forma de saber e
tentar importa-las - que e o que este ficheiro faz.

Codigo de saida: 0 = tudo bem, 1 = falta alguma coisa.
"""
import importlib
import sys

# O que o codigo importa mesmo. `fitz` e o nome de import do PyMuPDF e
# `dateutil` vem do pandas - por isso nao chega olhar para o requirements.txt.
NECESSARIAS = [
    ('fitz', 'PyMuPDF - leitura das tabelas do PDF'),
    ('pandas', 'pandas - tabelas e Excel/CSV'),
    ('numpy', 'numpy - vem com o pandas'),
    ('openpyxl', 'openpyxl - escrita dos .xlsx'),
    ('dateutil', 'python-dateutil - vem com o pandas'),
    ('requests', 'requests - cliente do AMALIA (so para a comparacao de criterios)'),
]

# Sem estas o projeto NAO corre. A ausencia das outras e um aviso.
ESSENCIAIS = {'fitz', 'pandas', 'numpy', 'openpyxl', 'dateutil'}


def main() -> int:
    faltam, avisos = [], []
    for modulo, descricao in NECESSARIAS:
        try:
            importlib.import_module(modulo)
        except Exception as exc:
            (faltam if modulo in ESSENCIAIS else avisos).append((modulo, descricao, exc))

    if not faltam and not avisos:
        print(f"       todas as {len(NECESSARIAS)} bibliotecas OK.")
        return 0

    for modulo, descricao, exc in avisos:
        print(f"       AVISO: falta {modulo} ({descricao})")
        print(f"              {type(exc).__name__}: {exc}")
    if avisos and not faltam:
        print("       O essencial esta ca - a extracao corre.")
        return 0

    print()
    for modulo, descricao, exc in faltam:
        print(f"       EM FALTA: {modulo}  ({descricao})")
        print(f"                 {type(exc).__name__}: {exc}")
    print()
    print(f"       Python usado: {sys.executable}")
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
