# -*- coding: utf-8 -*-
"""Where is data missing across the project? One Excel file, in English.

WHY THIS IS NOT JUST "COUNT THE BLANKS"
---------------------------------------
Several columns are empty for a whole regime BY DESIGN — a TUA certificate has
no `Meio Recetor` and an LURH has no `Sentido da decisao`. Counting those as
missing produces a report nobody trusts. So completeness is measured PER
REGIME, and a field that is simply not part of a regime is labelled as such
instead of being reported as a gap.

SHEETS
------
  1. Summary              one line per issue, with counts
  2. Licence field gaps   one line per licence, listing what it is missing
  3. Field completeness   every column x regime, % filled, gap or by-design
  4. No limits or monitoring   licences carrying no VLE and/or no autocontrolo
  5. PDFs never extracted      files in the archive with no extraction
  6. Extracted not in master   extractions missing from master_lurh
  7. Other tables         completeness of the remaining dashboard tables

Usage:
    python missing_data_report.py [--out <file.xlsx>]
"""
from __future__ import annotations

import datetime
import glob
import json
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
FEED_DIR = os.path.join(ROOT, 'data', 'powerbi')
LICENSES = os.path.join(FEED_DIR, 'Licenses.xlsx')
PDFS = os.path.join(ROOT, 'data', 'pdfs')
OUT = os.path.join(ROOT, 'outputs', 'reports', 'Missing_Data_Report.xlsx')

# Fields that only exist in one regime. Blank in the other is correct, not a gap.
LURH_ONLY = {'Meio Recetor', 'Denominação do meio recetor', 'REGIÃO', 'ARH', 'Município'}
TUA_ONLY: set = set()

# What a licence must have to be usable in the compliance dashboard.
ESSENTIAL = ['Nº TUA', 'Estabelecimento', 'Data de Entrada em Vigor',
             'Data de Validade', 'CODMAXIMO', 'Latitude', 'Longitude',
             'Nível de tratamento']


def _filled(series):
    return series.astype(str).str.strip() != ''


def _sub(regime, col):
    if col in LURH_ONLY:
        return regime == 'LURH'
    if col in TUA_ONLY:
        return regime == 'TUA'
    return True


def field_completeness(d):
    rows = []
    for c in d.columns:
        for regime in ('TUA', 'LURH'):
            s = d[d['TUA/LURH'] == regime]
            if not len(s):
                continue
            applies = _sub(regime, c)
            n = int(_filled(s[c]).sum())
            pct = round(100.0 * n / len(s), 1)
            rows.append({
                'Table': 'Licenses', 'Column': c, 'Regime': regime,
                'Rows': len(s), 'Filled': n, 'Filled %': pct,
                'Applies to this regime': 'yes' if applies else 'no (by design)',
                'Verdict': ('not applicable' if not applies
                            else 'complete' if pct >= 99
                            else 'minor gap' if pct >= 90
                            else 'GAP'),
            })
    return pd.DataFrame(rows)


def licence_gaps(d):
    rows = []
    for _, r in d.iterrows():
        regime = r['TUA/LURH']
        miss = [c for c in ESSENTIAL
                if c in d.columns and _sub(regime, c) and not str(r[c]).strip()]
        if not miss:
            continue
        rows.append({
            'Licence': r['Nº TUA'], 'Regime': regime,
            'ETAR': r['Estabelecimento'], 'File': r['file_name'],
            'Missing fields': ', '.join(miss), 'How many': len(miss),
        })
    return pd.DataFrame(rows).sort_values('How many', ascending=False) if rows else pd.DataFrame()


def no_limits_or_monitoring():
    """Licences whose extraction found no VLE table and/or no monitoring rows."""
    rows = []
    for folder, key in (('extraction_JSON_TUA', 'Nº TUA'),
                        ('extraction_JSON_LURH', 'Nº Licença')):
        for f in sorted(glob.glob(os.path.join(ROOT, 'src', folder, '*_extracted.json'))):
            try:
                with open(f, encoding='utf-8') as fh:
                    d = json.load(fh)
            except Exception:
                continue
            g = d.get('dados_gerais', {})
            lid = str(g.get(key, '')).strip()
            cond = d.get('condicoes_rejeicao') or d.get('condicoes_descarga') or []
            auto = d.get('autocontrolo')
            nauto = len(auto.get('rows', [])) if isinstance(auto, dict) else len(auto or [])
            if cond and nauto:
                continue
            rows.append({
                'Licence': lid or '(no licence number)',
                'Regime': 'TUA' if 'TUA' in folder else 'LURH',
                'File': d.get('file_name', os.path.basename(f)),
                'Discharge limits (VLE)': len(cond),
                'Monitoring rows': nauto,
                'Problem': ('no limits AND no monitoring' if not cond and not nauto
                            else 'no discharge limits' if not cond
                            else 'no monitoring rows'),
            })
    return pd.DataFrame(rows)


def never_extracted():
    """Archive PDFs with no matching extraction (matched on a normalised name)."""
    def slug(n):
        return ''.join(ch for ch in n.lower() if ch.isalnum())

    done = set()
    for folder in ('extraction_JSON_TUA', 'extraction_JSON_LURH'):
        for f in glob.glob(os.path.join(ROOT, 'src', folder, '*_extracted.json')):
            try:
                with open(f, encoding='utf-8') as fh:
                    done.add(slug(json.load(fh).get('file_name', '')))
            except Exception:
                pass
    rows = []
    for p in sorted(glob.glob(os.path.join(PDFS, '**', '*.pdf'), recursive=True)):
        if slug(os.path.basename(p)) in done:
            continue
        rel = os.path.relpath(p, PDFS).replace(os.sep, '/')
        rows.append({'File': os.path.basename(p), 'Folder': os.path.dirname(rel),
                     'Size (MB)': round(os.path.getsize(p) / 1e6, 2)})
    return pd.DataFrame(rows)


def extracted_not_in_master():
    m = os.path.join(ROOT, 'src', 'Extraction_Code_LURH', 'master_lurh.xlsx')
    if not os.path.isfile(m):
        return pd.DataFrame()
    have = set(pd.read_excel(m, sheet_name='Licenses', dtype=str)
               .fillna('')['Nº Licença'].str.strip())
    rows, seen = [], set()
    for f in sorted(glob.glob(os.path.join(ROOT, 'src', 'extraction_JSON_LURH', '*_extracted.json'))):
        try:
            with open(f, encoding='utf-8') as fh:
                d = json.load(fh)
        except Exception:
            continue
        lid = str(d.get('dados_gerais', {}).get('Nº Licença', '')).strip()
        if not lid or lid in seen:
            continue
        seen.add(lid)
        if lid not in have:
            rows.append({'Licence': lid, 'File': d.get('file_name', '')})
    return pd.DataFrame(rows)


def other_tables():
    rows = []
    for p in sorted(glob.glob(os.path.join(FEED_DIR, '*.xlsx'))):
        name = os.path.basename(p)
        if name == 'Licenses.xlsx':
            continue
        try:
            d = pd.read_excel(p, dtype=str).fillna('')
        except Exception:
            rows.append({'Table': name, 'Column': '(unreadable)', 'Rows': 0,
                         'Filled': 0, 'Filled %': 0.0, 'Verdict': 'UNREADABLE'})
            continue
        for c in d.columns:
            n = int(_filled(d[c]).sum())
            pct = round(100.0 * n / len(d), 1) if len(d) else 0.0
            rows.append({'Table': name, 'Column': str(c)[:60], 'Rows': len(d),
                         'Filled': n, 'Filled %': pct,
                         'Verdict': ('EMPTY COLUMN' if n == 0
                                     else 'complete' if pct >= 99
                                     else 'minor gap' if pct >= 90 else 'GAP')})
    return pd.DataFrame(rows)


def main(argv) -> int:
    out = argv[argv.index('--out') + 1] if '--out' in argv else OUT
    if not os.path.isfile(LICENSES):
        print(f'ERROR: feed not found:\n  {LICENSES}')
        return 1

    lic = pd.read_excel(LICENSES, dtype=str).fillna('')
    comp = field_completeness(lic)
    gaps = licence_gaps(lic)
    nolim = no_limits_or_monitoring()
    noext = never_extracted()
    notmaster = extracted_not_in_master()
    others = other_tables()

    real_gap_cols = comp[comp['Verdict'] == 'GAP']
    summary = pd.DataFrame([
        {'Area': 'Licences in the dashboard', 'Finding': 'total rows', 'Count': len(lic)},
        {'Area': 'Licence fields', 'Finding': 'licences missing >=1 essential field',
         'Count': len(gaps)},
        {'Area': 'Licence fields', 'Finding': 'column/regime combinations under 90% filled',
         'Count': len(real_gap_cols)},
        {'Area': 'Compliance content', 'Finding': 'licences with no discharge limits and/or no monitoring',
         'Count': len(nolim)},
        {'Area': 'Compliance content', 'Finding': 'of those, with NEITHER',
         'Count': int((nolim['Problem'] == 'no limits AND no monitoring').sum()) if len(nolim) else 0},
        {'Area': 'Coverage', 'Finding': 'archive PDFs never extracted', 'Count': len(noext)},
        {'Area': 'Coverage', 'Finding': 'extracted LURH licences absent from master_lurh',
         'Count': len(notmaster)},
        {'Area': 'Other dashboard tables', 'Finding': 'columns that are entirely empty',
         'Count': int((others['Verdict'] == 'EMPTY COLUMN').sum()) if len(others) else 0},
        {'Area': 'Report', 'Finding': 'generated', 'Count': datetime.date.today().isoformat()},
    ])

    os.makedirs(os.path.dirname(out), exist_ok=True)
    with pd.ExcelWriter(out, engine='openpyxl') as xl:
        summary.to_excel(xl, sheet_name='1 Summary', index=False)
        (gaps if len(gaps) else pd.DataFrame({'(none)': []})).to_excel(
            xl, sheet_name='2 Licence field gaps', index=False)
        comp.to_excel(xl, sheet_name='3 Field completeness', index=False)
        (nolim if len(nolim) else pd.DataFrame({'(none)': []})).to_excel(
            xl, sheet_name='4 No limits or monitoring', index=False)
        (noext if len(noext) else pd.DataFrame({'(none)': []})).to_excel(
            xl, sheet_name='5 PDFs never extracted', index=False)
        (notmaster if len(notmaster) else pd.DataFrame({'(none)': []})).to_excel(
            xl, sheet_name='6 Extracted not in master', index=False)
        others.to_excel(xl, sheet_name='7 Other tables', index=False)

    print(summary.to_string(index=False))
    print()
    print(f'  written to: {out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv[1:]))
