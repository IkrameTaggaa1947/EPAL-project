# -*- coding: utf-8 -*-
"""Uma linha por ETAR: que licenças tem, e se cobrem o mesmo período.

Responde a três perguntas de uma vez:

  * a ETAR tem TUA, LURH, ou as duas?
  * quando tem as duas, os períodos coincidem, sobrepõem-se, ou são
    completamente separados?
  * o que está em vigor HOJE?

A ETAR é identificada pelo CODMAXIMO (o código do activo), não pelo nome: o
mesmo sítio aparece escrito de maneiras diferentes nos documentos ("ETAR
Fundão" / "ETAR do Fundão"), e contar por nome inflaciona o número de ETAR.
As licenças sem CODMAXIMO não pertencem a ETAR nenhuma e vão para uma folha à
parte, para não desaparecerem sem se dar por isso.

Quando uma ETAR tem mais do que uma licença do mesmo regime, a comparação usa
a MAIS RECENTE de cada um, e a coluna "Outras licenças" diz quantas ficaram de
fora.

Uso:
    python relatorio_tua_vs_lurh.py [--out <ficheiro.xlsx>]
"""
from __future__ import annotations

import datetime
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
FEED = os.path.join(ROOT, 'data', 'powerbi', 'Licenses.xlsx')
SAIDA = os.path.join(ROOT, 'outputs', 'reports', 'ETAR_TUA_vs_LURH.xlsx')

HOJE = datetime.date.today()


def _data(v):
    v = str(v).strip()
    for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.datetime.strptime(v, fmt).date()
        except ValueError:
            pass
    return None


def _estado(fim):
    if fim is None:
        return 'sem data'
    return 'Em vigor' if fim >= HOJE else 'Caducada'


def _comparar(ti, tf, li, lf):
    """Como se relacionam os dois períodos."""
    if None in (ti, tf, li, lf):
        return 'Sem datas suficientes'
    if ti == li and tf == lf:
        return 'Mesmo período'
    if ti <= lf and li <= tf:
        return 'Sobrepõem-se (períodos diferentes)'
    return 'Períodos separados (não se tocam)'


def construir(feed_path=FEED) -> tuple:
    f = pd.read_excel(feed_path, dtype=str).fillna('')
    f['cod'] = f['CODMAXIMO'].str.strip()
    f['ini'] = f['Data de Entrada em Vigor'].map(_data)
    f['fim'] = f['Data de Validade'].map(_data)

    sem_cod = f[f['cod'] == ''].copy()
    com = f[f['cod'] != '']

    linhas = []
    for cod, s in com.groupby('cod'):
        tua = s[s['TUA/LURH'] == 'TUA'].sort_values('fim', na_position='first')
        lur = s[s['TUA/LURH'] == 'LURH'].sort_values('fim', na_position='first')
        nome = ''
        for v in list(s['Estabelecimento']) + list(s['Designação']):
            if str(v).strip():
                nome = str(v).strip()
                break

        t = tua.iloc[-1] if len(tua) else None
        l = lur.iloc[-1] if len(lur) else None
        extra = max(len(tua) - 1, 0) + max(len(lur) - 1, 0)

        if t is not None and l is not None:
            regime = 'TUA + LURH'
            periodo = _comparar(t['ini'], t['fim'], l['ini'], l['fim'])
            et, el = _estado(t['fim']), _estado(l['fim'])
            # 'sem data' NAO e o mesmo que 'caducada': nao se sabe. Dizer que
            # a ETAR nao tem nada em vigor por falta de data seria inventar.
            if et == 'Em vigor' and el == 'Em vigor':
                situacao = 'As duas em vigor'
            elif et == 'Em vigor':
                situacao = 'Só a TUA em vigor'
            elif el == 'Em vigor':
                situacao = 'Só a LURH em vigor'
            elif 'sem data' in (et, el):
                situacao = 'Por confirmar (falta data de validade)'
            else:
                situacao = 'NENHUMA em vigor'
        elif t is not None:
            regime, periodo = 'Só TUA', 'n.a. (só tem uma licença)'
            e = _estado(t['fim'])
            situacao = {'Em vigor': 'TUA em vigor',
                        'Caducada': 'TUA caducada — nada em vigor',
                        'sem data': 'Por confirmar (falta data de validade)'}[e]
        else:
            regime, periodo = 'Só LURH', 'n.a. (só tem uma licença)'
            e = _estado(l['fim'])
            situacao = {'Em vigor': 'LURH em vigor',
                        'Caducada': 'LURH caducada — nada em vigor',
                        'sem data': 'Por confirmar (falta data de validade)'}[e]

        linhas.append({
            'CODMAXIMO': cod,
            'ETAR': nome,
            'Regimes': regime,
            'Mesmo período?': periodo,
            'Situação hoje': situacao,
            'TUA nº': t['Nº TUA'] if t is not None else '',
            'TUA início': t['ini'] if t is not None else None,
            'TUA fim': t['fim'] if t is not None else None,
            'TUA estado': _estado(t['fim']) if t is not None else '',
            'LURH nº': l['Nº TUA'] if l is not None else '',
            'LURH início': l['ini'] if l is not None else None,
            'LURH fim': l['fim'] if l is not None else None,
            'LURH estado': _estado(l['fim']) if l is not None else '',
            'Outras licenças': extra,
        })

    tab = pd.DataFrame(linhas).sort_values(['Regimes', 'ETAR'])
    fora = sem_cod[['Nº TUA', 'TUA/LURH', 'Estabelecimento',
                    'Data de Entrada em Vigor', 'Data de Validade', 'file_name']]
    return tab, fora


def main(argv) -> int:
    saida = SAIDA
    if '--out' in argv:
        saida = argv[argv.index('--out') + 1]
    if not os.path.isfile(FEED):
        print(f'ERRO: nao encontrei o feed:\n  {FEED}')
        return 1

    tab, fora = construir()
    os.makedirs(os.path.dirname(saida), exist_ok=True)
    with pd.ExcelWriter(saida, engine='openpyxl') as xl:
        tab.to_excel(xl, sheet_name='Por ETAR', index=False)
        fora.to_excel(xl, sheet_name='Sem CODMAXIMO', index=False)
        (tab.groupby(['Regimes', 'Mesmo período?']).size()
            .rename('ETAR').reset_index()
            .to_excel(xl, sheet_name='Resumo', index=False))

    print(f'  ETAR no relatorio : {len(tab)}')
    print(f'  licencas de fora  : {len(fora)} (sem CODMAXIMO)')
    print()
    print('  POR REGIME:')
    for reg, n in tab['Regimes'].value_counts().items():
        print(f'     {reg:14} {n:4}')
    print()
    print('  DAS QUE TEM AS DUAS — os periodos coincidem?')
    duas = tab[tab['Regimes'] == 'TUA + LURH']
    for k, n in duas['Mesmo período?'].value_counts().items():
        print(f'     {k:38} {n:4}')
    print()
    print('  SITUACAO HOJE:')
    for k, n in tab['Situação hoje'].value_counts().items():
        print(f'     {k:34} {n:4}')
    print()
    print(f'  Gravado em: {saida}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
