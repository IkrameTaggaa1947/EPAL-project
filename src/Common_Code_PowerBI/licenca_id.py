# -*- coding: utf-8 -*-
"""Quando e que duas licencas sao A MESMA licenca?

PORQUE NAO CHEGA O Nº TUA
-------------------------
Uma RENOVACAO mantem o mesmo Nº TUA e muda so o periodo. A ETAR Cardigos tem
duas licencas com o numero TUA20230130000334:

    27-01-2023 a 26-01-2026   (original)
    27-01-2026 a 26-01-2031   (renovacao)

Enquanto a chave do master foi o Nº TUA sozinho, a segunda APAGAVA a primeira:
ficava uma linha so, com o ficheiro novo e as datas antigas -- e quem procurava
"Cardigos" via a licenca velha.

O QUE IDENTIFICA UMA LICENCA
----------------------------
A combinacao  DESIGNACAO + DATA DE ENTRADA EM VIGOR + DATA DE VALIDADE.
Duas linhas com a mesma combinacao sao a mesma licenca (uma re-extracao do
mesmo documento); qualquer diferenca faz delas licencas diferentes, e as duas
ficam no master.

O Nº TUA NAO entra de proposito: um documento reemitido pode trazer numero novo
para o mesmo periodo, e isso continua a ser a mesma licenca.
"""
from __future__ import annotations

import re
import unicodedata

COLUNA = 'Id Versão'


def _norm(s) -> str:
    """Minusculas, sem acentos, sem espacos a mais. Nada mais agressivo:
    'ETAR Cardigos' e 'ETAR de Cardigos' SAO sitios diferentes."""
    s = unicodedata.normalize('NFKD', str(s or ''))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip().lower()


def _data(s) -> str:
    """Datas comparaveis venham como vierem: 27-01-2026, 27/01/2026, 27.01.2026."""
    m = re.search(r'(\d{2})[-/.](\d{2})[-/.](\d{4})', str(s or ''))
    return '%s-%s-%s' % m.groups() if m else ''


def identidade(designacao, entrada_em_vigor, validade) -> str:
    """A chave de uma licenca. Legivel de proposito: quando aparece numa folha
    de Excel tem de se perceber logo de que licenca se trata."""
    return '%s|%s|%s' % (_norm(designacao), _data(entrada_em_vigor), _data(validade))


def de_extracao(d: dict) -> str:
    """A identidade a partir de uma extracao, venha ela de que regime vier.

    Os dois extractores nao guardam isto no mesmo sitio:

      TUA   as datas vivem em 'enquadramento' (Data de Entrada em Vigor /
            Data de Validade) e o nome em dados_gerais.Estabelecimento;
      LURH  nao tem 'enquadramento' nenhum -- as datas estao em dados_gerais
            como 'Data de Início' / 'Data de Validade', e o nome ou vem de
            rejeicao['Designação da rejeição'] ou do Requerente.

    Enquanto isto so olhou para a forma do TUA, TODAS as licencas LURH davam a
    identidade vazia '||' -- ou seja, todas iguais entre si.
    """
    g = d.get('dados_gerais') or {}
    e = d.get('enquadramento') or {}
    rej = d.get('rejeicao') or {}

    designacao = (g.get('Estabelecimento') or g.get('Designação')
                  or g.get('Designacao') or rej.get('Designação da rejeição')
                  or g.get('Requerente') or '')
    vigor = (e.get('Data de Entrada em Vigor') or g.get('Data de Entrada em Vigor')
             or g.get('Data de Início') or g.get('Data de Inicio') or '')
    validade = (e.get('Data de Validade') or g.get('Data de Validade') or '')
    return identidade(designacao, vigor, validade)


def carimbar(sheets: dict, ident: str) -> None:
    """Poe COLUNA em TODAS as linhas de TODAS as folhas desta licenca.

    As folhas-filho (Conditions, Autocontrolo, ...) so tinham o Nº TUA, por isso
    nao havia como distinguir as condicoes da licenca velha das da nova. Com
    esta coluna passa a haver."""
    for rows in sheets.values():
        for r in rows:
            if isinstance(r, dict):
                r[COLUNA] = ident


# ------------------------------------------------------------------ self-check
def _demo():
    a = identidade('ETAR Cardigos', '27-01-2023', '26-01-2026')
    b = identidade('ETAR Cardigos', '27-01-2026', '26-01-2031')
    assert a != b, 'a renovacao tem de ser uma licenca DIFERENTE'

    # a mesma licenca escrita de outra maneira continua a ser a mesma
    assert a == identidade('  etar   CARDIGOS ', '27/01/2023', '26.01.2026')

    # sitios diferentes nunca colidem
    assert a != identidade('ETAR Vale de Cardigos', '27-01-2023', '26-01-2026')

    # sem datas ainda da uma chave utilizavel (a designacao sozinha)
    assert identidade('ETAR X', '', '') == 'etar x||'

    # a partir de uma extracao TUA (datas em 'enquadramento')
    d = {'dados_gerais': {'Estabelecimento': 'ETAR Cardigos'},
         'enquadramento': {'Data de Entrada em Vigor': '27-01-2026',
                           'Data de Validade': '26-01-2031'}}
    assert de_extracao(d) == b

    # e a partir de uma LURH, que nao tem 'enquadramento' e usa 'Data de Inicio'
    lurh = {'dados_gerais': {'Requerente': 'Aguas do Vale do Tejo, S.A.',
                             'Data de Início': '17-10-2016',
                             'Data de Validade': '16-10-2026'},
            'rejeicao': {'Designação da rejeição': 'ETAR de Évora'}}
    got = de_extracao(lurh)
    assert got == identidade('ETAR de Évora', '17-10-2016', '16-10-2026'), got
    assert got != '||', 'uma LURH nunca pode dar identidade vazia'

    # carimbar chega a todas as folhas
    sheets = {'Licenses': [{'Nº TUA': 'X'}], 'Conditions': [{'Nº TUA': 'X'}, {'Nº TUA': 'X'}]}
    carimbar(sheets, b)
    assert all(r[COLUNA] == b for rows in sheets.values() for r in rows)
    print('licenca_id self-check OK')


if __name__ == '__main__':
    _demo()
