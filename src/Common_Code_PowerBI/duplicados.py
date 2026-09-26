# -*- coding: utf-8 -*-
"""Nao voltar a extrair uma licenca que ja esta no arquivo.

DUAS PERGUNTAS DIFERENTES
-------------------------
1. "Este PDF ja esta no arquivo?"  -> compara-se o CONTEUDO (SHA-256).
   Resposta segura: ou os bytes sao iguais ou nao sao. Custa 0,01 s e apanha o
   caso comum -- alguem poe o mesmo ficheiro outra vez, com o mesmo nome ou com
   outro. Nesses casos NAO se extrai nada.

2. "Esta licenca (ETAR + periodo) ja esta no arquivo?"  -> so se sabe ao certo
   DEPOIS de extrair, e por isso a decisao de saltar fica para depois.

PORQUE NAO SE DECIDE (2) ANTES DE EXTRAIR
-----------------------------------------
Seria bom: a extracao custa ~45 s e a leitura do texto custa 0,04 s. Mediu-se,
nos 402 PDFs do arquivo, a fiabilidade de adivinhar o periodo sem extrair:

    datas lidas do texto do documento .....  33% certas
    datas lidas do nome do ficheiro .......  94% certas
    sem datas nenhumas ....................  17% dos ficheiros

Os 6% errados do nome nao sao inofensivos: o nome diz uma coisa e o documento
diz outra (a ETAR Aranhas tem '23.02.2023 a 22.02.2028' no nome e
'13-05-2022 a 12-05-2027' no documento). Saltar por causa disso seria PERDER uma
licenca verdadeira, em silencio. Perder 45 s e barato; perder uma licenca nao.

Por isso: salta-se sem extrair APENAS quando o conteudo e identico. Quando a
licenca ja existe mas o ficheiro e outro, extrai-se e DIZ-SE -- o operador fica
a saber que ha duas versoes do mesmo documento no arquivo.
"""
from __future__ import annotations

import glob
import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import licenca_id                                          # noqa: E402


def sha256(caminho: str, bloco: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(caminho, 'rb') as fh:
        for pedaco in iter(lambda: fh.read(bloco), b''):
            h.update(pedaco)
    return h.hexdigest()


def indice_do_arquivo(pasta_pdfs: str) -> dict:
    """{sha256: caminho} de tudo o que ja esta arquivado."""
    idx = {}
    for p in glob.glob(os.path.join(pasta_pdfs, '**', '*.pdf'), recursive=True):
        try:
            idx.setdefault(sha256(p), p)
        except OSError:
            pass
    return idx


def ja_arquivado(pdf: str, indice: dict) -> str:
    """O caminho do gemeo no arquivo, ou '' se este PDF e novo."""
    try:
        return indice.get(sha256(pdf), '')
    except OSError:
        return ''


def identidades_conhecidas(*pastas_json) -> dict:
    """{identidade: file_name} do que ja foi extraido."""
    import json

    out = {}
    for pasta in pastas_json:
        for f in glob.glob(os.path.join(pasta, '*_extracted.json')):
            try:
                with open(f, encoding='utf-8') as fh:
                    d = json.load(fh)
            except Exception:
                continue
            ident = licenca_id.de_extracao(d)
            if ident and ident != '||':
                out.setdefault(ident, d.get('file_name', os.path.basename(f)))
    return out


# ------------------------------------------------------------------ self-check
def _demo():
    import tempfile

    tmp = tempfile.mkdtemp()
    arquivo = os.path.join(tmp, 'pdfs', 'NEW', 'TUA')
    os.makedirs(arquivo)

    a = os.path.join(arquivo, 'licenca.pdf')
    open(a, 'wb').write(b'%PDF-1.4 conteudo da licenca')
    idx = indice_do_arquivo(os.path.join(tmp, 'pdfs'))

    # mesmo conteudo, OUTRO nome -> e o mesmo documento
    b = os.path.join(tmp, 'outro nome.pdf')
    open(b, 'wb').write(b'%PDF-1.4 conteudo da licenca')
    assert ja_arquivado(b, idx) == a, 'devia reconhecer o gemeo pelo conteudo'

    # conteudo diferente -> e novo, mesmo que o nome seja igual
    c = os.path.join(tmp, 'licenca.pdf')
    open(c, 'wb').write(b'%PDF-1.4 outra versao')
    assert ja_arquivado(c, idx) == '', 'bytes diferentes nao sao duplicado'

    import shutil
    shutil.rmtree(tmp, ignore_errors=True)
    print('duplicados self-check OK')


if __name__ == '__main__':
    _demo()
