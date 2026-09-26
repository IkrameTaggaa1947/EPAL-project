# -*- coding: utf-8 -*-
"""Todas as licenças do arquivo chegaram ao dashboard?

Segue cada PDF de data/pdfs pelas quatro etapas e diz onde se perdeu:

    PDF  ->  JSON extraído  ->  master  ->  feed do Power BI

A ligação PDF->JSON é feita pelo nome do ficheiro (é o que o extrator grava
em `file_name`), e daí para a frente pelo número da licença. Contar ficheiros
não chega: o mesmo PDF pode ter sido extraído duas vezes com nomes
diferentes, e um JSON pode existir sem nunca ter entrado no master.

Uso:
    python src\\_verificar_cobertura.py            # resumo
    python src\\_verificar_cobertura.py --listar   # + nomes do que falta
"""
from __future__ import annotations

import glob
import json
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

REGIMES = {
    'TUA': {
        'pastas': [os.path.join('data', 'pdfs', 'OLD', 'TUA'),
                   os.path.join('data', 'pdfs', 'NEW', 'TUA')],
        'json': os.path.join('src', 'extraction_JSON_TUA'),
        'master': os.path.join('src', 'Extraction_Code_TUA', 'master_tua.xlsx'),
        'chave': 'Nº TUA',
    },
    'LURH': {
        'pastas': [os.path.join('data', 'pdfs', 'OLD', 'LURH'),
                   os.path.join('data', 'pdfs', 'NEW', 'LURH')],
        'json': os.path.join('src', 'extraction_JSON_LURH'),
        'master': os.path.join('src', 'Extraction_Code_LURH', 'master_lurh.xlsx'),
        'chave': 'Nº Licença',
    },
}

FEED = os.path.join('data', 'powerbi', 'Licenses.xlsx')


def _slug(nome):
    """Nome comparável: só letras e dígitos, minúsculas.

    O mesmo PDF foi extraído com nomes diferentes ao longo do tempo
    ('Ade_Lic_L020555.2019.RH3 (03.01.2020 a 02.01.2025).pdf' e
    'Ade_Lic_L020555_2019_RH3_03_01_2020_a_02_01_2025.pdf'), por isso comparar
    o nome tal e qual dá falsos 'nunca extraído'."""
    return ''.join(c for c in nome.lower() if c.isalnum())


def _carregar_json(pasta):
    """{nome_do_pdf: chave}, {slug_do_nome: chave} dos JSON extraídos."""
    por_ficheiro, por_slug = {}, {}
    for p in glob.glob(os.path.join(ROOT, pasta, '*_extracted.json')):
        try:
            with open(p, encoding='utf-8') as fh:
                d = json.load(fh)
        except Exception:
            continue
        nome = d.get('file_name', '')
        g = d.get('dados_gerais', {})
        chave = str(g.get('Nº TUA') or g.get('Nº Licença') or '').strip()
        if nome:
            por_ficheiro[nome] = chave
            por_slug[_slug(nome)] = chave
    return por_ficheiro, por_slug


def _chave_do_nome(nome, regime):
    """Número da licença lido do próprio nome do ficheiro, quando dá.

    Usa o parser que o projeto já tem (qa.from_filename), que é a mesma
    verdade que o QA usa para a verificação cruzada."""
    pasta = 'Extraction_Code_TUA' if regime == 'TUA' else 'Extraction_Code_LURH'
    caminho = os.path.join(HERE, pasta)
    if caminho not in sys.path:
        sys.path.insert(0, caminho)
    try:
        import importlib
        qa = importlib.import_module('qa')
        got = qa.from_filename(nome)
    except Exception:
        return ''
    if not got:
        return ''
    return str(got.get('Nº TUA') or got.get('Nº Licença') or '').strip()


def _chaves_master(caminho, chave):
    if not os.path.isfile(os.path.join(ROOT, caminho)):
        return set()
    d = pd.read_excel(os.path.join(ROOT, caminho), sheet_name='Licenses')
    if chave not in d.columns:
        return set()
    return {str(v).strip() for v in d[chave].dropna() if str(v).strip()}


def _chaves_feed():
    p = os.path.join(ROOT, FEED)
    if not os.path.isfile(p):
        return set()
    d = pd.read_excel(p)
    col = 'Nº TUA' if 'Nº TUA' in d.columns else d.columns[0]
    return {str(v).strip() for v in d[col].dropna() if str(v).strip()}


def main(argv):
    listar = '--listar' in argv
    feed = _chaves_feed()
    problemas = 0

    for regime, cfg in REGIMES.items():
        print()
        print('=' * 68)
        print(f' {regime}')
        print('=' * 68)
        por_ficheiro, por_slug = _carregar_json(cfg['json'])
        master = _chaves_master(cfg['master'], cfg['chave'])

        sem_extracao, sem_master, sem_feed, ok = [], [], [], 0
        total = 0
        for pasta in cfg['pastas']:
            pdfs = sorted(glob.glob(os.path.join(ROOT, pasta, '*.pdf')))
            for p in pdfs:
                total += 1
                nome = os.path.basename(p)
                etiqueta = f"{os.path.basename(os.path.dirname(p))}/{nome}"

                # três formas de ligar o PDF à sua extração, da mais fiável
                # para a menos: nome exacto, nome normalizado, e o número da
                # licença lido do próprio nome do ficheiro.
                chave = por_ficheiro.get(nome)
                if chave is None:
                    chave = por_slug.get(_slug(nome))
                do_nome = _chave_do_nome(nome, regime)
                if chave is None and do_nome and do_nome in master:
                    chave = do_nome        # está no master, com outro nome

                if chave is None:
                    sem_extracao.append(etiqueta + (f'  [{do_nome}]' if do_nome else ''))
                elif not chave:
                    sem_master.append(etiqueta + '  (extraído sem número de licença)')
                elif chave not in master:
                    sem_master.append(f'{etiqueta}  [{chave}]')
                elif chave not in feed:
                    sem_feed.append(f'{etiqueta}  [{chave}]')
                else:
                    ok += 1

        print(f'  PDFs no arquivo          : {total}')
        print(f'  chegaram ao Power BI     : {ok}')
        print(f'  NUNCA extraídos          : {len(sem_extracao)}')
        print(f'  extraídos, fora do master: {len(sem_master)}')
        print(f'  no master, fora do feed  : {len(sem_feed)}')
        problemas += len(sem_extracao) + len(sem_master) + len(sem_feed)

        for titulo, lista in (('NUNCA extraídos', sem_extracao),
                              ('extraídos mas fora do master', sem_master),
                              ('no master mas fora do feed', sem_feed)):
            if not lista:
                continue
            print()
            print(f'  --- {titulo} ({len(lista)})')
            mostrar = lista if listar else lista[:8]
            for x in mostrar:
                print(f'      {x}')
            if not listar and len(lista) > 8:
                print(f'      ... e mais {len(lista) - 8}  (correr com --listar)')

        # JSON sem PDF no arquivo: extracções órfãs
        slugs_pdf = set()
        for pasta in cfg['pastas']:
            slugs_pdf |= {_slug(os.path.basename(p))
                          for p in glob.glob(os.path.join(ROOT, pasta, '*.pdf'))}
        orfaos = sorted(n for n in por_ficheiro if _slug(n) not in slugs_pdf)
        if orfaos:
            print()
            print(f'  --- JSON sem PDF no arquivo ({len(orfaos)})')
            print('      (extracções antigas, ou o PDF foi movido/renomeado)')
            for x in (orfaos if listar else orfaos[:5]):
                print(f'      {x}')
            if not listar and len(orfaos) > 5:
                print(f'      ... e mais {len(orfaos) - 5}')

    print()
    print('=' * 68)
    if problemas == 0:
        print(' Todas as licenças do arquivo estão no dashboard.')
        return 0
    print(f' {problemas} licença(s) não chegaram ao dashboard.')
    return 1


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
