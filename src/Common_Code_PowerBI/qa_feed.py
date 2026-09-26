# -*- coding: utf-8 -*-
"""Quality gate on the Power BI feed — run after the tables are rebuilt.

A badly extracted PDF used to reach the dashboard silently: a parameter cut in
half ("pH (Escala de"), a date that no longer parses, a licence with no
parameter at all. This script inspects data/powerbi/*.xlsx and reports every
anomaly it can prove, so a bad extraction is caught before anyone reads it.

Exit code 0 = clean or warnings only, 1 = at least one ERROR.
Run:  python src/qa_feed.py [--strict]
      --strict turns warnings into errors (use it in the pipeline once the
      known data gaps are filled).
"""
import datetime
import os
import re
import sys
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
ROOT = os.path.dirname(os.path.dirname(HERE))   # project root (this module is in src/Common_Code_PowerBI)
FEED = os.getenv('EPAL_POWERBI_OUT', os.path.join(ROOT, 'data', 'powerbi'))

from param_names import abbreviate            # noqa: E402

# A parameter name is suspicious when it is cut mid-word: the extractor lost the
# rest of the line. These are the shapes seen in real broken extractions.
TRUNCATED = re.compile(r'\($|\bde$|\be$|\bda$|\bdo$|,$|-$')


class Report:
    def __init__(self):
        self.errors = []
        self.warnings = []
        self.stats = []

    def error(self, table, msg):
        self.errors.append(f'{table}: {msg}')

    def warn(self, table, msg):
        self.warnings.append(f'{table}: {msg}')

    def stat(self, msg):
        self.stats.append(msg)


def rows(name):
    import openpyxl
    path = os.path.join(FEED, f'{name}.xlsx')
    if not os.path.exists(path):
        return None
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    it = wb.active.values
    try:
        hdr = [str(c).strip() if c is not None else '' for c in next(it)]
    except StopIteration:
        wb.close()
        return []
    out = [dict(zip(hdr, r)) for r in it]
    wb.close()
    return out


def parse_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        for f in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
            try:
                return datetime.datetime.strptime(v[:10], f).date()
            except ValueError:
                pass
    return None


def check_licences(rep):
    data = rows('Licenses')
    if data is None:
        return rep.error('Licenses', 'ficheiro em falta')
    rep.stat(f'Licenses: {len(data)} licenças')

    ids = [str(r.get('Nº TUA') or '').strip() for r in data]
    blank = sum(1 for i in ids if not i)
    if blank:
        rep.error('Licenses', f'{blank} linha(s) sem Nº TUA')
    dups = [k for k, n in Counter(i for i in ids if i).items() if n > 1]
    if dups:
        rep.error('Licenses', f'Nº TUA repetido: {dups[:5]}')

    no_date = [r['Nº TUA'] for r in data
               if not parse_date(r.get('Validade (ISO)')) and not parse_date(r.get('Data de Validade'))]
    if no_date:
        rep.warn('Licenses', f'{len(no_date)} licença(s) sem data de validade legível')

    # a validity far outside the plausible range means a parsing mistake
    today = datetime.date.today()
    for r in data:
        d = parse_date(r.get('Validade (ISO)'))
        if d and not (today.year - 30 <= d.year <= today.year + 30):
            rep.error('Licenses', f"{r.get('Nº TUA')}: validade implausível ({d})")

    for col in ('Longitude', 'Latitude'):
        bad = []
        for r in data:
            v = str(r.get(col) or '').replace(',', '.').strip()
            if not v:
                continue
            try:
                x = float(v)
            except ValueError:
                bad.append(r.get('Nº TUA'))
                continue
            lo, hi = (-32, 0) if col == 'Longitude' else (30, 45)   # Portugal + Açores
            if not lo <= x <= hi:
                bad.append(f"{r.get('Nº TUA')}={x}")
        if bad:
            rep.error('Licenses', f'{col} fora de Portugal ou ilegível: {bad[:5]}')
    return data


def check_parameters(rep, licences):
    known = {str(r.get('Nº TUA') or '').strip() for r in (licences or [])}
    for name, idcol in (('Conditions', 'Nº TUA'), ('Autocontrolo', 'Nº TUA'),
                        ('Licencas_Criterios', 'IdLicenca')):
        data = rows(name)
        if data is None:
            rep.error(name, 'ficheiro em falta')
            continue
        rep.stat(f'{name}: {len(data)} linhas')
        if not data:
            rep.error(name, 'tabela vazia')
            continue

        orphans = {str(r.get(idcol) or '').strip() for r in data} - known - {''}
        if orphans:
            rep.error(name, f'{len(orphans)} licença(s) inexistente(s) em Licenses: '
                            f'{sorted(orphans)[:3]}')

        raw_col = 'Parâmetro' if 'Parâmetro' in data[0] else 'Parâmetro (documento)'
        cut = [str(r.get(raw_col)) for r in data
               if r.get(raw_col) and TRUNCATED.search(str(r[raw_col]).strip())]
        if cut:
            rep.error(name, f'{len(cut)} parâmetro(s) truncado(s) pela extração: '
                            f'{sorted(set(cut))[:3]}')

        empty = sum(1 for r in data if not str(r.get('Parametro') or '').strip())
        if empty:
            rep.warn(name, f'{empty} linha(s) sem abreviatura de parâmetro')

        # an abbreviation that is really a full sentence means the table in
        # param_names.py has no entry for it
        longs = {str(r.get('Parametro')) for r in data
                 if len(str(r.get('Parametro') or '')) > 20}
        if longs:
            rep.warn(name, f'parâmetro(s) sem abreviatura conhecida: {sorted(longs)[:3]}')


def check_criteria(rep):
    bridge = rows('Licencas_Criterios')
    catalog = rows('Catalogo_Criterios')
    if bridge is None or catalog is None:
        return
    ids = {str(c.get('IdCriterio') or '').strip() for c in catalog}
    missing = {str(b.get('IdCriterio') or '').strip() for b in bridge} - ids - {''}
    if missing:
        rep.error('Licencas_Criterios',
                  f'{len(missing)} IdCriterio ausente(s) do catálogo: {sorted(missing)[:5]}')
    no_crit = sum(1 for b in bridge if not str(b.get('IdCriterio') or '').strip())
    if no_crit:
        rep.warn('Licencas_Criterios', f'{no_crit} linha(s) sem critério associado')


def check_renewal(rep):
    data = rows('Renovacao_Proposta')
    if data is None:
        return
    per = defaultdict(list)
    for r in data:
        per[str(r.get('Nº TUA') or '')].append(r)
    rep.stat(f'Renovacao_Proposta: {len(per)} licenças a renovar, {len(data)} campos')
    sizes = {len(v) for v in per.values()}
    if len(sizes) > 1:
        rep.error('Renovacao_Proposta',
                  f'número de campos difere entre licenças: {sorted(sizes)}')
    for tua, fields in per.items():
        filled = sum(1 for f in fields if str(f.get('Valor') or '').strip() != 'Em falta')
        if filled == 0:
            rep.error('Renovacao_Proposta', f'{tua}: nenhum campo preenchido')


def main(argv):
    strict = '--strict' in argv
    rep = Report()
    if not os.path.isdir(FEED):
        print(f'! feed folder not found: {FEED}', file=sys.stderr)
        return 1

    licences = check_licences(rep)
    check_parameters(rep, licences)
    check_criteria(rep)
    check_renewal(rep)

    print('QA do feed Power BI —', FEED)
    for s in rep.stats:
        print('  ', s)
    if rep.warnings:
        print(f'\nAVISOS ({len(rep.warnings)}):')
        for w in rep.warnings:
            print('   -', w)
    if rep.errors:
        print(f'\nERROS ({len(rep.errors)}):')
        for e in rep.errors:
            print('   -', e)
    if not rep.errors and not rep.warnings:
        print('\ntudo consistente')

    return 1 if rep.errors or (strict and rep.warnings) else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
