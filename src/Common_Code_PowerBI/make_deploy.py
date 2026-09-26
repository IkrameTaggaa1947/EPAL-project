# -*- coding: utf-8 -*-
"""Build a deployment package: the dashboard and its data, nothing else.

The semantic model reads a handful of Excel/CSV files and nothing else — no
Python, no R, no database. So whoever only READS the dashboard needs just those
files; the extraction code stays with whoever maintains the pipeline.

The list is read off the model's own partitions (see needed()), never hardcoded:
a hardcoded list had already drifted out of sync and produced packages that
could not refresh.

    deploy/
    ├── EPAL_Licencas_Dashboard/     the PBIP project
    ├── data/                        the files the model reads
    ├── LEIA-ME.txt                  how to open it, where the data comes from
    └── MANIFEST.txt                 what was packaged, with sizes and date

The DataFolder parameter is rewritten to the data/ folder of the package, so
the project opens on any machine once the folder is copied.

Run:  python src/Common_Code_PowerBI/make_deploy.py [--out <folder>]
      python src/Common_Code_PowerBI/make_deploy.py --list   (so mostrar o que o modelo le)
"""
import datetime
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))   # project root (this module is in src/Common_Code_PowerBI)
PROJECT = os.path.join(ROOT, 'reports', 'EPAL_Licencas_Dashboard')
FEED = os.path.join(ROOT, 'data', 'powerbi')

MODEL_DEF = os.path.join(PROJECT, 'EPAL_Licencas_Dashboard.SemanticModel', 'definition')

# Every way a partition can name a file under DataFolder:
#   File.Contents(DataFolder & "\Licenses.xlsx")        -> direct
#   File.Contents(DataFolder & "\TUA\Sec (3.7).csv")    -> in a subfolder
#   Ler("Conditions.xlsx", "Conditions")                -> via a local M helper
_DIRECT = re.compile(r'DataFolder\s*&\s*"\\([^"]+\.(?:xlsx|csv))"')
_VIA_HELPER = re.compile(r'\(\s*"([^"\\]+\.(?:xlsx|csv))"\s*,')


def needed(model_def=MODEL_DEF):
    """The files the semantic model actually reads, read off the model itself.

    This used to be a hardcoded list with a 'keep in sync with the partitions'
    comment, and it had drifted: it asked for a table the model never reads and
    omitted five that it does, so the package it built could not refresh. Derive
    it instead, and it cannot drift again."""
    files = set()
    tables = os.path.join(model_def, 'tables')
    for name in sorted(os.listdir(tables)):
        if not name.endswith('.tmdl'):
            continue
        text = open(os.path.join(tables, name), encoding='utf-8').read()
        files.update(_DIRECT.findall(text))
        if 'DataFolder' in text:                 # only trust the helper form here
            files.update(_VIA_HELPER.findall(text))
    return sorted(files, key=str.lower)

LEIAME = """DASHBOARD DE LICENÇAS EPAL — pacote de distribuição
{date}

COMO ABRIR
  1. Copiar esta pasta inteira para o computador (manter os dois subdirectórios
     juntos: o projeto procura os dados em .\\data).
  2. Abrir EPAL_Licencas_Dashboard\\EPAL_Licencas_Dashboard.pbip
     com o Power BI Desktop.
  3. Início -> Atualizar.

DE ONDE VÊM OS DADOS
  Os ficheiros em data\\ são GERADOS pelo pipeline de extração a partir das
  licenças em PDF. Não devem ser editados à mão: são reescritos a cada
  execução do pipeline.

  Para ter dados novos, é preciso que alguém execute o pipeline e volte a
  copiar a pasta data\\ — ou publicar o relatório no Power BI Service, onde a
  atualização pode ser agendada.

VISUAIS PERSONALIZADOS (IMPORTANTE)
  O relatório usa dois visuais do AppSource em 6 páginas:
      calendarHeatmap...   textFilter...
  O Power BI Desktop vai buscá-los ao AppSource ao abrir o ficheiro. Num
  computador sem internet, ou numa organização que bloqueie o AppSource,
  essas páginas ficam em branco. Não são ficheiros que este pacote possa
  transportar — têm de ser permitidos pela política da organização.

SE OS DADOS NÃO CARREGAM
  O projeto guarda o caminho da pasta de dados no parâmetro "DataFolder".
  Este pacote já o aponta para a pasta data\\ ao lado do relatório. Se a pasta
  for movida: Power BI -> Transformar dados -> Gerir parâmetros -> DataFolder.

CONTEÚDO
  EPAL_Licencas_Dashboard\\   relatório e modelo (PBIP)
  data\\                      {n} tabelas em Excel lidas pelo modelo
"""


def human(n):
    return f'{n/1024:.0f} KB' if n < 1024 ** 2 else f'{n/1024**2:.1f} MB'


def _demo():
    """The parser must see both reference styles the model actually uses."""
    found = needed()
    assert found, 'nenhum ficheiro encontrado nas particoes do modelo'
    for must in ('Licenses.xlsx', 'Conditions.xlsx', 'Autocontrolo.xlsx',
                 'Licencas_Criterios.xlsx'):
        assert must in found, (must, found)
    assert any(f.endswith('.csv') and ('\\' in f or '/' in f) for f in found), \
        'as particoes em subpasta (data/powerbi/TUA/*.csv) nao foram apanhadas'
    print(f'make_deploy self-check OK ({len(found)} ficheiros lidos do modelo)')


def main(argv):
    if '--list' in argv:
        for n in needed():
            print(n)
        return 0
    if '--self-check' in argv:
        _demo()
        return 0
    out = ROOT
    if '--out' in argv:
        out = argv[argv.index('--out') + 1]
    dest = os.path.join(out, 'deploy')

    wanted = needed()
    missing = [n for n in wanted if not os.path.exists(os.path.join(FEED, n))]
    if missing:
        print(f'! tabelas em falta em {FEED}: {missing}', file=sys.stderr)
        print('  correr primeiro:  python src/Common_Code_PowerBI/make_powerbi_all.py && '
              'python src/Common_Code_PowerBI/make_powerbi_extras.py', file=sys.stderr)
        return 1

    if os.path.isdir(dest):
        shutil.rmtree(dest)
    os.makedirs(os.path.join(dest, 'data'))

    shutil.copytree(PROJECT, os.path.join(dest, 'EPAL_Licencas_Dashboard'),
                    ignore=shutil.ignore_patterns('.pbi', 'cache.abf', '*.tmp'))

    lines = []
    for name in wanted:
        src = os.path.join(FEED, name)
        out_path = os.path.join(dest, 'data', name)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)   # names may carry a subfolder
        shutil.copy2(src, out_path)
        lines.append(f'  data/{name.replace(os.sep, "/")}  {human(os.path.getsize(src))}')

    # point DataFolder at the data/ folder shipped next to the report
    expr = os.path.join(dest, 'EPAL_Licencas_Dashboard',
                        'EPAL_Licencas_Dashboard.SemanticModel', 'definition',
                        'expressions.tmdl')
    target = os.path.abspath(os.path.join(dest, 'data')).replace('/', '\\')
    s = open(expr, encoding='utf-8').read()
    s = re.sub(r'(expression DataFolder = )"[^"]*"', lambda m: m.group(1) + '"' + target + '"', s)
    open(expr, 'w', encoding='utf-8', newline='').write(s)

    today = datetime.date.today().strftime('%d/%m/%Y')
    open(os.path.join(dest, 'LEIA-ME.txt'), 'w', encoding='utf-8').write(
        LEIAME.format(date=today, n=len(wanted)))
    total = sum(os.path.getsize(os.path.join(dp, f))
                for dp, _, fs in os.walk(dest) for f in fs)
    open(os.path.join(dest, 'MANIFEST.txt'), 'w', encoding='utf-8').write(
        f'Pacote gerado em {today}\nTotal: {human(total)}\n\n'
        + 'EPAL_Licencas_Dashboard/  (projeto PBIP)\n' + '\n'.join(lines) + '\n')

    if os.name != 'nt':
        print('! ATENÇÃO: gerado fora do Windows — o caminho DataFolder acima não é\n'
              '  válido no Power BI Desktop. Voltar a correr este script no Windows\n'
              '  (python src\\make_deploy.py) antes de distribuir o pacote.',
              file=sys.stderr)
    print(f'pacote pronto: {dest}  ({human(total)})')
    print(f'  DataFolder -> {target}')
    for l in lines:
        print(l)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
