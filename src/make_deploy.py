# -*- coding: utf-8 -*-
"""Build a deployment package: the dashboard and its data, nothing else.

The semantic model reads eight Excel files and nothing else — no Python, no R,
no database. So whoever only READS the dashboard needs just those files; the
extraction code stays with whoever maintains the pipeline.

    deploy/
    ├── EPAL_Licencas_Dashboard/     the PBIP project
    ├── data/                        the eight .xlsx the model reads
    ├── LEIA-ME.txt                  how to open it, where the data comes from
    └── MANIFEST.txt                 what was packaged, with sizes and date

The DataFolder parameter is rewritten to the data/ folder of the package, so
the project opens on any machine once the folder is copied.

Run:  python src/make_deploy.py [--out <folder>]
"""
import datetime
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PROJECT = os.path.join(ROOT, 'reports', 'EPAL_Licencas_Dashboard')
FEED = os.path.join(ROOT, 'data', 'powerbi')

# exactly what the model reads — keep in sync with the partitions
NEEDED = ['Licenses', 'Conditions', 'Autocontrolo', 'Avaliacao',
          'Catalogo_Criterios', 'Licencas_Criterios', 'Condicoes_Prioritarias',
          'Renovacao_Proposta']

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


def main(argv):
    out = ROOT
    if '--out' in argv:
        out = argv[argv.index('--out') + 1]
    dest = os.path.join(out, 'deploy')

    missing = [n for n in NEEDED if not os.path.exists(os.path.join(FEED, f'{n}.xlsx'))]
    if missing:
        print(f'! tabelas em falta em {FEED}: {missing}', file=sys.stderr)
        print('  correr primeiro:  python src/make_powerbi_all.py && '
              'python src/make_powerbi_extras.py', file=sys.stderr)
        return 1

    if os.path.isdir(dest):
        shutil.rmtree(dest)
    os.makedirs(os.path.join(dest, 'data'))

    shutil.copytree(PROJECT, os.path.join(dest, 'EPAL_Licencas_Dashboard'),
                    ignore=shutil.ignore_patterns('.pbi', 'cache.abf', '*.tmp'))

    lines = []
    for name in NEEDED:
        src = os.path.join(FEED, f'{name}.xlsx')
        shutil.copy2(src, os.path.join(dest, 'data', f'{name}.xlsx'))
        lines.append(f'  data/{name}.xlsx  {human(os.path.getsize(src))}')

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
        LEIAME.format(date=today, n=len(NEEDED)))
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
