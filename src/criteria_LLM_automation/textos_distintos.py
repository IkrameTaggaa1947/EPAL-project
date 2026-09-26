# -*- coding: utf-8 -*-
"""Perguntar ao AMALIA UMA vez por texto legal, nao uma vez por linha.

PORQUE
------
A pergunta semantica -- 'estes dois textos legais querem dizer o mesmo?' -- nao
depende da ETAR nem do parametro: depende so do PAR DE TEXTOS. E os textos
repetem-se muito:

    1810 linhas em Conditions
     129 pares de textos distintos          <- 14x menos perguntas

Nesta maquina o modelo gera 0,8 tokens/s (9B em CPU, sem GPU). Uma resposta de
~500 tokens leva perto de 10 minutos. A diferenca e concreta:

    uma pergunta por linha   ~314 horas
    uma pergunta por par      ~22 horas

A resposta fica em CACHE no disco, por isso uma segunda passagem nao volta a
pagar nada, e uma execucao interrompida retoma onde ia.

NORMALIZACAO
------------
Dois textos que so diferem num ponto final, num 'n.º' escrito de outra maneira
ou num espaco a mais SAO o mesmo texto. Sem isso pagavam-se respostas repetidas
(142 pares em vez de 129).
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata

CACHE_PADRAO = os.path.join(os.getenv('LOCALAPPDATA') or os.path.expanduser('~'),
                            'EPAL', 'amalia_cache.json')


def normalizar(s) -> str:
    """A forma comparavel de um texto legal."""
    s = unicodedata.normalize('NFKD', str(s or ''))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'\s+', ' ', s).strip().lower()
    s = re.sub(r'n\.?\s*[oº°]\s*', 'n', s)        # n.º / nº / n o -> n
    s = re.sub(r'\s*([.,;:])\s*', r'\1', s)       # espacos a volta da pontuacao
    return re.sub(r'[\s.,;:]+$', '', s)           # pontuacao final


def chave(leg_antiga, aval_antiga, leg_nova, aval_nova) -> str:
    """Identificador curto e estavel do par de textos."""
    cru = '\x1f'.join(normalizar(x) for x in
                      (leg_antiga, aval_antiga, leg_nova, aval_nova))
    return hashlib.sha1(cru.encode('utf-8')).hexdigest()[:16]


def pares_distintos(linhas) -> dict:
    """{chave: (exemplo_da_linha, quantas_linhas)} — o trabalho REAL a fazer.

    `linhas` sao dicts com 'leg_o'/'aval_o'/'leg_n'/'aval_n'."""
    fora = {}
    for r in linhas:
        k = chave(r.get('leg_o'), r.get('aval_o'), r.get('leg_n'), r.get('aval_n'))
        if k in fora:
            fora[k] = (fora[k][0], fora[k][1] + 1)
        else:
            fora[k] = (r, 1)
    return fora


class Cache:
    """As respostas ja dadas, guardadas no disco entre execucoes."""

    def __init__(self, caminho=None):
        self.caminho = caminho or CACHE_PADRAO
        self.dados = {}
        if os.path.isfile(self.caminho):
            try:
                with open(self.caminho, encoding='utf-8') as fh:
                    self.dados = json.load(fh)
            except Exception:
                self.dados = {}          # cache ilegivel nunca trava o trabalho

    def get(self, k):
        return self.dados.get(k)

    def put(self, k, valor):
        self.dados[k] = valor
        self.gravar()                    # gravar a cada resposta: se parar a
                                         # meio, nao se perde o que ja custou

    def gravar(self):
        os.makedirs(os.path.dirname(self.caminho), exist_ok=True)
        tmp = self.caminho + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as fh:
            json.dump(self.dados, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, self.caminho)


def aplicar(linhas, cache: Cache, campos) -> int:
    """Poe em cada linha a resposta guardada para o seu par. Devolve quantas
    linhas ficaram preenchidas."""
    n = 0
    for r in linhas:
        v = cache.get(chave(r.get('leg_o'), r.get('aval_o'),
                            r.get('leg_n'), r.get('aval_n')))
        if v:
            for c in campos:
                r[c] = v.get(c, '')
            n += 1
    return n


# ------------------------------------------------------------------ self-check
def _demo():
    import tempfile

    # textos que so diferem em pontuacao/n.º SAO o mesmo
    a = 'Anexo XVIII do Decreto-Lei n.º 236/98, de 1 de agosto.'
    b = 'Anexo XVIII do Decreto-Lei nº 236/98 , de 1 de agosto'
    assert normalizar(a) == normalizar(b), (normalizar(a), normalizar(b))
    assert chave(a, 'x', 'y', 'z') == chave(b, 'x', 'y', 'z')

    # textos diferentes continuam diferentes
    assert chave(a, 'x', 'y', 'z') != chave('Anexo I do DL 152/97', 'x', 'y', 'z')

    # o agrupamento conta as repeticoes
    linhas = [{'leg_o': a, 'aval_o': 'q', 'leg_n': 'n', 'aval_n': 'w'},
              {'leg_o': b, 'aval_o': 'q', 'leg_n': 'n', 'aval_n': 'w'},
              {'leg_o': 'outro', 'aval_o': 'q', 'leg_n': 'n', 'aval_n': 'w'}]
    d = pares_distintos(linhas)
    assert len(d) == 2, d
    assert sorted(n for _, n in d.values()) == [1, 2]

    # a cache sobrevive ao disco e aplica-se a TODAS as linhas do par
    with tempfile.TemporaryDirectory() as tmp:
        c = Cache(os.path.join(tmp, 'c.json'))
        k = chave(a, 'q', 'n', 'w')
        c.put(k, {'am_muda': 'Não', 'am_just': 'mesma norma'})
        c2 = Cache(os.path.join(tmp, 'c.json'))          # relido do disco
        assert c2.get(k)['am_muda'] == 'Não'
        n = aplicar(linhas, c2, ('am_muda', 'am_just'))
        assert n == 2, n                                  # as duas do mesmo par
        assert linhas[0]['am_muda'] == 'Não'
        assert 'am_muda' not in linhas[2]
    print('textos_distintos self-check OK')


if __name__ == '__main__':
    _demo()
