# -*- coding: utf-8 -*-
"""A identidade de uma licenca SEM fazer a extracao toda.

PARA QUE SERVE
--------------
Uma licenca e identificada por ETAR + periodo de validade (ver licenca_id.py).
Ate agora essa identidade so se sabia DEPOIS de extrair, o que e absurdo: se a
licenca ja esta no arquivo, gastaram-se 45 segundos para deitar fora o
resultado -- e ninguem foi avisado de que era repetida.

Os tres campos de que a identidade precisa estao no TEXTO da primeira pagina.
Ler o texto custa 0,04 s; a extracao completa custa 45 s. Da para decidir mil
vezes antes de valer a pena extrair uma.

COMO LE AS DATAS
----------------
No texto, os rotulos vem todos juntos e so depois os valores:

    Data de Emissão
    Data de Entrada em Vigor
    Data de Validade
    04-12-2025
    27-01-2026
    26-01-2031

Por isso procura-se o ULTIMO rotulo 'Data de Validade' e leem-se as tres datas
seguintes. O ULTIMO de proposito: numa renovacao o documento traz primeiro o
periodo original e depois o novo, e o que vale e o novo.

Se o texto nao chegar, cai para as datas do NOME do ficheiro -- muitos trazem-nas.
"""
from __future__ import annotations

import os
import re
import sys

# O runtime\python corre em modo ISOLADO (ficheiro ._pth) e nao acrescenta a
# pasta do script ao sys.path. Sem isto, correr este ficheiro sozinho falha.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import licenca_id                                          # noqa: E402

DATA = re.compile(r'\b(\d{2})[-/.](\d{2})[-/.](\d{4})\b')
ROTULO_VALIDADE = re.compile(r'data\s+de\s+validade', re.I)


def _datas_do_nome(nome: str) -> list:
    """Datas plausiveis no nome do ficheiro, pela ordem em que aparecem."""
    out = []
    for d, m, a in DATA.findall(nome or ''):
        if 1 <= int(d) <= 31 and 1 <= int(m) <= 12 and 1900 <= int(a) <= 2100:
            out.append('%s-%s-%s' % (d, m, a))
    return out


# O nome da ETAR nao se chama o mesmo nos dois regimes: o TUA diz
# 'Estabelecimento', a LURH diz 'Designação da rejeição'. Sem a segunda, todas
# as licencas LURH davam nome vazio.
ROTULOS_NOME = ('estabelecimento', 'designacao da rejeicao',
                'designação da rejeição', 'designacao da rejeição')


def _sem_acentos(s: str) -> str:
    import unicodedata
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if not unicodedata.combining(c)).lower()


def _estabelecimento(linhas) -> str:
    for i, l in enumerate(linhas):
        chave = _sem_acentos(l)
        if any(chave.startswith(_sem_acentos(r)) for r in ROTULOS_NOME):
            resto = l.split(':', 1)[-1].strip()
            if resto and _sem_acentos(resto) not in [_sem_acentos(r) for r in ROTULOS_NOME]:
                return resto
            return linhas[i + 1] if i + 1 < len(linhas) else ''
    return ''


def _periodo(texto: str) -> tuple:
    """(entrada em vigor, validade) a partir do bloco de datas mais recente."""
    marcas = list(ROTULO_VALIDADE.finditer(texto))
    if marcas:
        depois = texto[marcas[-1].end():]
        datas = ['%s-%s-%s' % g for g in DATA.findall(depois)[:3]]
        if len(datas) >= 3:                # emissao, entrada em vigor, validade
            return datas[1], datas[2]
        if len(datas) == 2:                # sem data de emissao
            return datas[0], datas[1]
    return '', ''


def ler(pdf_path: str, paginas: int = 3) -> dict:
    """{'estabelecimento','vigor','validade','identidade','fonte'} — barato."""
    import fitz

    try:
        doc = fitz.open(pdf_path)
        texto = '\n'.join(doc[i].get_text() for i in range(min(paginas, doc.page_count)))
        doc.close()
    except Exception:
        texto = ''

    linhas = [l.strip() for l in texto.splitlines() if l.strip()]
    est = _estabelecimento(linhas)
    vigor, validade = _periodo(texto)
    fonte = 'texto'

    if not (vigor and validade):           # o documento nao deu: tentar o nome
        do_nome = _datas_do_nome(os.path.basename(pdf_path))
        if len(do_nome) >= 2:
            vigor, validade = do_nome[0], do_nome[-1]
            fonte = 'nome do ficheiro'
        else:
            fonte = 'incompleta'

    return {'estabelecimento': est, 'vigor': vigor, 'validade': validade,
            'fonte': fonte,
            'identidade': licenca_id.identidade(est, vigor, validade)}


def conhecidas(*pastas_json) -> set:
    """As identidades que ja existem, lidas das extracoes que estao no disco."""
    import glob
    import json

    vistas = set()
    for pasta in pastas_json:
        for f in glob.glob(os.path.join(pasta, '*_extracted.json')):
            try:
                with open(f, encoding='utf-8') as fh:
                    vistas.add(licenca_id.de_extracao(json.load(fh)))
            except Exception:
                pass
    vistas.discard('||')
    return vistas


# ------------------------------------------------------------------ self-check
def _demo():
    linhas = ['Estabelecimento', 'ETAR Cardigos', 'Data de Emissão',
              'Data de Entrada em Vigor', 'Data de Validade']
    assert _estabelecimento(linhas) == 'ETAR Cardigos'
    assert _estabelecimento(['Estabelecimento: ETAR X']) == 'ETAR X'

    # renovacao: dois blocos, vale o ULTIMO
    txt = ('Data de Emissão Data de Entrada em Vigor Data de Validade '
           '30-01-2023 27-01-2023 26-01-2026 '
           'Data de Emissão Data de Entrada em Vigor Data de Validade '
           '04-12-2025 27-01-2026 26-01-2031')
    assert _periodo(txt) == ('27-01-2026', '26-01-2031'), _periodo(txt)

    # sem rotulo -> nada (e depois cai para o nome do ficheiro)
    assert _periodo('texto qualquer sem datas') == ('', '')

    # nome do ficheiro: aceita os dois estilos, e nao confunde codigos com datas
    assert _datas_do_nome('Ade TUA20240731002347 (02.01.2025 a 01.01.2030).pdf') == \
        ['02-01-2025', '01-01-2030']
    assert _datas_do_nome('Fronteira_2012.000813.000.T.L.RJ.DAR.pdf') == []
    print('identidade_rapida self-check OK')


if __name__ == '__main__':
    _demo()
