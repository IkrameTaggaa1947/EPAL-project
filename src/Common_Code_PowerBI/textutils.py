# -*- coding: utf-8 -*-
"""Ajudantes partilhados pelos DOIS extratores (TUA e LURH).

PORQUE EXISTE ESTE FICHEIRO
---------------------------
extract_tua.py e extract_lurh.py tinham copias identicas destas funcoes. A
maior parte do que os dois ficheiros partilham e so o NOME - os formatos dos
documentos sao mesmo diferentes e o codigo tem de ser diferente. Mas estas
cinco eram copia a copia, e tres delas (`norm`, `_key`, `col_idx`) sao a
REGRA de correspondencia de colunas por nome: o coracao da leitura das
tabelas. Uma correcao a essa regra tinha de ser feita duas vezes, e foi
exatamente esse tipo de divergencia que deixou o TUA sem a correcao que o
LURH ja tinha (ver o historico de pipeline.py).

Quem precisar de comportamento diferente NAO deve editar aqui: deve manter a
sua propria versao no seu extrator, como o `clean_text` do LURH faz (o corpus
LURH usa \\x02 como hifen; o TUA nao).
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime

from openpyxl.utils import get_column_letter


def norm(s):
    """Texto comparavel: sem acentos, sem espacos repetidos, minusculas."""
    if not s:
        return ''
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().lower()


def _key(s):
    """norm() tambem sem espacos - para comparar cabecalhos de coluna."""
    return norm(s).replace(' ', '')


def col_idx(header, *names):
    """Indice da 1.a coluna cujo cabecalho CONTEM um dos nomes dados.

    Corresponder por NOME (e nao por posicao) e o que torna os extratores
    robustos a certificados com as colunas por outra ordem."""
    for i, c in enumerate(header):
        kc = _key(c)
        if kc and any(_key(n) in kc for n in names):
            return i
    return None


def _widths(ws, widths):
    """Larguras de coluna de uma folha Excel, da esquerda para a direita."""
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _status(validade):
    """'Em vigor' / 'Caducada' a partir de uma data dd-mm-aaaa. '' se nao ler."""
    try:
        d = datetime.strptime(validade, '%d-%m-%Y').date()
        return 'Em vigor' if d >= datetime.now().date() else 'Caducada'
    except Exception:
        return ''


# ------------------------------------------------------------------ self-check
def _demo():
    assert norm('  Águas   RESIDUAIS ') == 'aguas residuais'
    assert norm(None) == '' and norm('') == ''
    assert _key('Nº  Licença') == 'nolicenca'

    hdr = ['Parâmetro', 'VLE (mg/L)', 'Frequência de amostragem']
    assert col_idx(hdr, 'parametro') == 0
    assert col_idx(hdr, 'vle') == 1
    assert col_idx(hdr, 'frequencia') == 2
    assert col_idx(hdr, 'nao existe') is None
    # acentos e maiusculas do documento nao contam
    assert col_idx(['FREQUENCIA'], 'Frequência') == 0

    assert _status('01-01-2000') == 'Caducada'
    assert _status('01-01-2099') == 'Em vigor'
    assert _status('nao e uma data') == ''
    print('textutils self-check OK')


if __name__ == '__main__':
    _demo()
