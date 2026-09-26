# -*- coding: utf-8 -*-
"""Pôr no master_lurh as licenças que foram extraídas mas nunca lá entraram.

PORQUE É PRECISO
----------------
O feed do Power BI e o master são construídos de sítios DIFERENTES:

  * `make_powerbi_all.py` lê a pasta `src/extraction_JSON_LURH/` — ou seja,
    tudo o que alguma vez foi extraído;
  * `master_lurh.xlsx` é uma acumulação à parte, escrita pelo pipeline no fim
    de cada licença.

Quando uma licença é extraída mas o master não chega a ser escrito (ficheiro
aberto no Excel, execução interrompida, uma versão antiga do pipeline), o JSON
fica e o master não. A partir daí o dashboard mostra mais licenças do que o
master tem, e ninguém dá por isso — foi assim que apareceram 700 no dashboard
contra 684 nos master.

Este script fecha essa diferença: procura os JSON cujo Nº Licença não está no
master e acrescenta-os, usando o MESMO upsert do pipeline (update_master_many),
para as folhas ficarem exactamente com a mesma forma.

Uso:
    python sync_master_com_json.py            # mostra o que falta, não altera
    python sync_master_com_json.py --aplicar  # acrescenta ao master
"""
from __future__ import annotations

import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd                                            # noqa: E402

import extract_lurh as E                                       # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
JSON_DIR = os.path.join(os.path.dirname(HERE), 'extraction_JSON_LURH')
MASTER = os.getenv('EPAL_LURH_MASTER', os.path.join(HERE, 'master_lurh.xlsx'))


def chaves_do_master(caminho: str) -> set:
    if not os.path.isfile(caminho):
        return set()
    d = pd.read_excel(caminho, sheet_name='Licenses', dtype=str).fillna('')
    col = 'Nº Licença'
    if col not in d.columns:
        return set()
    return {str(v).strip() for v in d[col] if str(v).strip()}


def em_falta(caminho: str):
    """[(chave, dados)] dos JSON que o master não tem. O mais recente ganha."""
    ja = chaves_do_master(caminho)
    vistos, falta = set(), []
    for f in sorted(glob.glob(os.path.join(JSON_DIR, '*_extracted.json'))):
        try:
            with open(f, encoding='utf-8') as fh:
                d = json.load(fh)
        except Exception:
            continue
        lid = str(d.get('dados_gerais', {}).get('Nº Licença', '')).strip()
        if not lid or lid in vistos:
            continue                       # sem número, ou já tratado
        vistos.add(lid)
        if lid not in ja:
            falta.append((lid, d))
    return falta


def main(argv) -> int:
    aplicar = '--aplicar' in argv
    falta = em_falta(MASTER)

    print(f'  master   : {os.path.basename(MASTER)}  ({len(chaves_do_master(MASTER))} licencas)')
    print(f'  em falta : {len(falta)}')
    for lid, d in falta[:25]:
        print(f'     {lid:24} {d.get("file_name","")[:52]}')
    if len(falta) > 25:
        print(f'     ... e mais {len(falta) - 25}')

    if not falta:
        print('\n  O master ja tem tudo o que foi extraido. Nada a fazer.')
        return 0
    if not aplicar:
        print('\n  ENSAIO - nada foi alterado. Repita com --aplicar.')
        return 0

    alvo, contagens, adiado = E.update_master_many([d for _, d in falta], MASTER)
    if adiado:
        print(f'\n  O master esta aberto/bloqueado. Os dados ficaram em '
              f'{os.path.basename(alvo)};\n  feche o Excel e volte a correr '
              f'(a proxima execucao do pipeline integra-o sozinha).')
        return 1
    print(f'\n  ACRESCENTADAS {len(falta)} licenca(s). O master tem agora '
          f'{contagens.get("Licenses", "?")} linhas em Licenses.')
    print('  Correr agora:  python src\\Common_Code_PowerBI\\make_powerbi_all.py')
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
