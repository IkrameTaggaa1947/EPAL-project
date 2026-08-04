# -*- coding: utf-8 -*-
"""Build powerbi_csv_all/ — TUA + LURH licences in one set of Power BI CSVs.

Reads the TUA export (src/tua/powerbi_csv/*.csv) and appends the LURH
licences (src/lurh/*_extracted.json) mapped into the same schema, with
the 'TUA/LURH' column distinguishing the two licence types.
Run after either pipeline updates:  python extraction/make_powerbi_all.py
"""
import csv, glob, json, os, re, sys
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)   # param_names lives next to this file
TUA_CSV = os.path.join(os.path.dirname(HERE), 'data', 'powerbi', 'TUA')
LURH_DIR = os.path.join(HERE, 'lurh')
OUT = os.getenv('EPAL_POWERBI_OUT', os.path.join(os.path.dirname(HERE), 'data', 'powerbi'))
CRITERIA_XLSX = os.path.join(os.path.dirname(HERE), 'data', 'reference', 'Comparacao_Criterios_Conformidade_TODAS_v2.xlsx')
MASTER_LURH = os.path.join(os.path.dirname(HERE), 'data', 'reference', 'master_lurh.xlsx')
CENTROIDE_XLSX = os.path.join(os.path.dirname(HERE), 'data', 'reference', 'TabelaCentroide_ETAR.xlsx')
LURH_CSV = os.path.join(os.path.dirname(HERE), 'data', 'powerbi', 'LURH')

ARH_BY_RH = {'RH1': ('ARH do Norte', 'Norte'), 'RH2': ('ARH do Norte', 'Norte'),
             'RH3': ('ARH do Norte', 'Norte'), 'RH4': ('ARH do Centro', 'Centro'),
             'RH4A': ('ARH do Centro', 'Centro'), 'RH5': ('ARH do Tejo e Oeste', 'Tejo e Oeste'),
             'RH5A': ('ARH do Tejo e Oeste', 'Tejo e Oeste'), 'RH6': ('ARH do Alentejo', 'Alentejo'),
             'RH7': ('ARH do Alentejo', 'Alentejo'), 'RH8': ('ARH do Algarve', 'Algarve')}

CRIT_COLS = ['Gama de valores (intervalo)', '≤ 100% VLE (dobro)', '≤ 150% VLE',
             '≤ uma ordem de grandeza do VLE', 'média mensal ≤ VLE', 'média anual ≤ VLE',
             'Quadro III DL 152/97 (borlas)']
CRIT_EXTRA = ['IdCritério correspondente', 'Estado do critério']

PARAM_COLS = ['Parametro', 'Unidade', 'Periodo']

# Column order of the files handed to Power BI: identity first, then what the
# parameter IS, then its limits, then how it is sampled, then the legal text.
COLUMN_ORDER = {
    'Conditions': [
        'Nº TUA', 'Parametro', 'Unidade', 'Periodo', 'Parâmetro',
        'VLE', 'VLE mín', 'VLE máx', 'VLE (% mín. redução)',
        'Carga máx. admissível (kg/dia)',
        'Frequência de amostragem', 'Tipo de amostragem',
        'Legislação aplicável', 'Avaliação da Conformidade Legal', 'Observações',
    ],
    'Autocontrolo': [
        'Nº TUA', 'Código', 'Parametro', 'Unidade', 'Periodo', 'Parâmetro',
        'Local de amostragem', 'Frequência de amostragem', 'Tipo de amostragem',
        'Nº análises requeridas', 'Nota tipo de amostragem', 'Observações',
    ],
    'Avaliacao': ['Nº TUA', 'Ordem', 'Código', 'Avaliação da conformidade'],
    'Legislacao': ['Nº TUA', 'Código', 'Legislação aplicável'],
}

# Criteria flags belong to Catalogo_Criterios only — a single source of truth.
# Conditions keeps the licence's own values, not the interpretation.
DROP_FROM_CONDITIONS = CRIT_COLS + CRIT_EXTRA + ['ParamKey', 'Chave']


def order_columns(stem, hdr):
    """Reorder a header to the canonical layout; unlisted columns keep their
    relative order at the end, so a new extracted field is never lost."""
    wanted = [c for c in COLUMN_ORDER.get(stem, []) if c in hdr]
    return wanted + [c for c in hdr if c not in wanted]


def number_avaliacao(rows):
    """'Código' arrives either as a sequence (1,2,3) or as a clause id (T000010)
    depending on the certificate. Add 'Ordem' — the position inside the licence —
    so every row can be sorted the same way, and keep the raw code as printed."""
    seen = {}
    for r in rows:
        lic = str(r.get('Nº TUA') or '')
        seen[lic] = seen.get(lic, 0) + 1
        r['Ordem'] = seen[lic]


def with_param_columns(hdr):
    """Header with Parametro/Unidade/Periodo right after 'Parâmetro'.

    'Parâmetro' (the raw text as printed in the licence) is kept as the audit
    trail; the dashboard uses the three derived columns.
    """
    hdr = [h for h in hdr if h not in PARAM_COLS]
    if 'Parâmetro' not in hdr:
        return hdr + PARAM_COLS
    at = hdr.index('Parâmetro') + 1
    return hdr[:at] + PARAM_COLS + hdr[at:]


def split_param_columns(rows):
    """Fill Parametro (abbreviation), Unidade and Periodo on every row."""
    from param_names import split_parameter
    for r in rows:
        short, unit, period = split_parameter(r.get('Parâmetro', ''))
        r['Parametro'], r['Unidade'], r['Periodo'] = short, unit, period


def clean_vle(value):
    from param_names import clean_vle as _clean
    return _clean(value)


def load_criteria():
    """(Nº TUA, Parâmetro catálogo) -> dict of the 7 criteria flags + match info,
    from the Criterios_0_1 sheet of the comparison workbook."""
    try:
        import openpyxl
        ws = openpyxl.load_workbook(CRITERIA_XLSX, read_only=True)['Criterios_0_1']
    except Exception:
        return {}
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h) for h in next(rows)]
    out = {}
    for r in rows:
        d = dict(zip(hdr, r))
        key = (str(d.get('Nº TUA') or ''), str(d.get('Parâmetro') or ''))
        rec = {c: d.get(c, '') for c in CRIT_COLS}
        rec['IdCritério correspondente'] = d.get('IdCritério correspondente', '')
        rec['Estado do critério'] = d.get('Estado', '')
        out[key] = rec
    return out

CAT_BY_KEY = {'pH': 'pH', 'CBO5': 'CBO5', 'CQO': 'CQO', 'SST': 'SST',
              'O&G': 'Óleo e gorduras', 'N': 'Nt', 'P': 'Pt', 'CF': 'CF', 'E.coli': 'E. coli'}

def lurh_criteria(text):
    """Interpret a LURH 'Avaliação da conformidade' text with the same rules
    used for the TUA criteria workbook (subset that occurs in LURH licences)."""
    import unicodedata
    t = unicodedata.normalize('NFKD', str(text or '')).encode('ascii', 'ignore').decode().lower()
    t = re.sub(r'\s+', ' ', t)
    f = {c: '' for c in CRIT_COLS}
    if len(t) < 6:
        return f, ''
    fired = False
    if re.search(r'artigo 69', t) and re.search(r'n\.?\s*o?\s*6', t.split('artigo')[0] + ' ' + t):
        f['média mensal ≤ VLE'] = 1; f['≤ 100% VLE (dobro)'] = 1; fired = True
    if 'alinea d' in t or ('nenhuma amostra excede' in t and '100' in t):
        f['≤ 100% VLE (dobro)'] = 1; fired = True
        if 'relacao estatistica' in t or 'quadro' in t or 'alinea d' in t:
            f['Quadro III DL 152/97 (borlas)'] = 1
    if 'anexo xviii' in t and not fired:
        return f, '⚠ Sem critério derivável (só legislação)'
    return f, ('derivado do texto (LURH)' if fired else '⚠ Sem critério derivável')

PARAMKEY = [('cbo', 'CBO5'), ('cqo', 'CQO'), ('sst', 'SST'), ('ph', 'pH'),
            ('gordur', 'O&G'), ('oleo', 'O&G'), ('azoto', 'N'), (' n ', 'N'),
            ('fosforo', 'P'), ('coliform', 'CF'), ('coli', 'E.coli')]

def paramkey(p):
    n = ' ' + re.sub(r'\s+', ' ', str(p or '')).lower() + ' '
    for sub, key in PARAMKEY:
        if sub in n:
            return key
    return (p or '').strip()[:12]

def vle_bounds(vle):
    s = str(vle or '').replace(',', '.')
    m = re.findall(r'\d+(?:\.\d+)?', s)
    if not m: return '', ''
    if len(m) >= 2 and re.search(r'-| a ', s): return m[0], m[1]
    return '', m[0]

def iso(d):
    m = re.match(r'(\d{2})-(\d{2})-(\d{4})', str(d or ''))
    return f'{m.group(3)}-{m.group(2)}-{m.group(1)}' if m else ''

def read_csv(name):
    p = os.path.join(TUA_CSV, name)
    with open(p, encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    return rows

INT_COLS = ('Ano de arranque', 'População servida à data',
            'Ano horizonte de projeto', 'População servida no horizonte',
            'Nº análises requeridas')
DEC_COLS = ('Longitude', 'Latitude')


def normalise_numbers(rows):
    """Make the numeric columns of Licenses.csv uniform across TUA and LURH.

    Two real defects this fixes, both introduced by reading LURH values through
    pandas and writing them with str():

      * integers arrived as floats — 'Ano de arranque' = "2002.0", which fails
        Int64.Type conversion in Power Query and turns the whole row into an
        error (this hit all 140 LURH rows);
      * coordinates used a DOT decimal separator on LURH rows while TUA rows
        used a COMMA. The Licenses query parses that column as "pt-PT", where a
        dot is a THOUSANDS separator, so "-7.536845" is not just an error risk,
        it silently becomes -7536845 and throws the pin off the map.

    Everything is written the way the TUA pipeline already writes it: integers
    without a decimal part, decimals with a comma.
    """
    def as_int(v):
        s = str(v or '').strip().replace(',', '.')
        if not s:
            return ''
        try:
            return str(int(round(float(s))))
        except ValueError:
            return ''

    def as_dec(v):
        s = str(v or '').strip()
        if not s:
            return ''
        try:
            return f'{float(s.replace(",", ".")):.6f}'.rstrip('0').rstrip('.').replace('.', ',')
        except ValueError:
            return ''

    for r in rows:
        for c in INT_COLS:
            if c in r:
                r[c] = as_int(r[c])
        for c in DEC_COLS:
            if c in r:
                r[c] = as_dec(r[c])


_CENTROIDE_CACHE = {}


def centroide_rows_and_keys(full=False):
    """Read TabelaCentroide_ETAR.xlsx (the ETAR asset registry) and add Chave_ETAR.

    Returns (header, set of keys, rows). Every column of the registry is kept —
    CODMAXIMO is what links an ETAR to WWTP_History, Centro_Operacional is what
    the Visão Geral slicer filters on, and the rest costs nothing to carry.
    """
    if 'rows' not in _CENTROIDE_CACHE:
        import pandas as pd
        from etar_key import chave_etar
        df = pd.read_excel(CENTROIDE_XLSX, dtype=str).fillna('')
        df['Chave_ETAR'] = df['DESIGNACAO_GNA'].map(chave_etar)
        df = df[df['Chave_ETAR'] != '']
        # the key is the ONE side of the relationship — it must not repeat
        dup = df['Chave_ETAR'][df['Chave_ETAR'].duplicated()].tolist()
        if dup:
            raise SystemExit(f'TabelaCentroide_ETAR: Chave_ETAR duplicada em {dup[:5]} — '
                             f'a relação com Licenses exige uma chave única.')
        _CENTROIDE_CACHE['hdr'] = list(df.columns)
        _CENTROIDE_CACHE['rows'] = df.to_dict('records')
        _CENTROIDE_CACHE['keys'] = set(df['Chave_ETAR'])
    if full:
        return _CENTROIDE_CACHE['hdr'], _CENTROIDE_CACHE['keys'], _CENTROIDE_CACHE['rows']
    return _CENTROIDE_CACHE['hdr'], _CENTROIDE_CACHE['keys']


def main():
    os.makedirs(OUT, exist_ok=True)
    lic_rows = read_csv('Licenses.csv')
    cond_rows = read_csv('Conditions.csv')
    lic_hdr = list(lic_rows[0].keys())
    # drop the legacy static_tables 'Critério 1..7' columns — superseded by the
    # workbook-derived criteria flags (verified per licence)
    LEGACY_CRIT = re.compile(r'^Critério \d')
    cond_hdr = [h for h in cond_rows[0].keys() if not LEGACY_CRIT.match(h)]
    for r in cond_rows:
        for h in list(r):
            if LEGACY_CRIT.match(h):
                del r[h]
    cond_hdr = cond_hdr + CRIT_COLS + CRIT_EXTRA
    today = date.today()
    criteria = load_criteria()
    # enrich existing TUA condition rows with the 7 criteria columns
    for r in cond_rows:
        cat = CAT_BY_KEY.get(r.get('ParamKey', ''), r.get('ParamKey', ''))
        rec = criteria.get((r.get('Nº TUA', ''), cat), {})
        for c in CRIT_COLS + CRIT_EXTRA:
            r[c] = rec.get(c, '')

    seen = set()
    # --- primary LURH source: master_lurh.xlsx (also exported to data/powerbi/LURH) ---
    if os.path.isfile(MASTER_LURH):
        import pandas as pd
        ml = pd.read_excel(MASTER_LURH, sheet_name=None)
        os.makedirs(LURH_CSV, exist_ok=True)
        for name, df in ml.items():
            if len(df):
                df.to_csv(os.path.join(LURH_CSV, f'{name}.csv'), index=False, encoding='utf-8-sig')
        mcond = {}
        for _, c in ml.get('Conditions', pd.DataFrame()).iterrows():
            mcond.setdefault(str(c.get('Nº Licença', '')), []).append(c)
        for _, g in ml['Licenses'].iterrows():
            lid = str(g.get('Nº Licença', '') or '')
            if not lid or lid in seen:
                continue
            seen.add(lid)
            val_iso = iso(g.get('Data de Validade'))
            try:
                vig = 'Em Vigor' if datetime.strptime(val_iso, '%Y-%m-%d').date() >= today else 'Caducada'
            except ValueError:
                vig = ''
            arh, reg = '', ''
            mrh = re.search(r'\.(RH\w+)$', lid)
            if mrh and mrh.group(1) in ARH_BY_RH:
                arh, reg = ARH_BY_RH[mrh.group(1)]
            def gv(col):
                v = g.get(col, '')
                return '' if pd.isna(v) else str(v)
            row = {k: '' for k in lic_hdr}
            row.update({
                'file_name': gv('file_name'), 'TUA/LURH': 'LURH', 'Nº TUA': lid,
                'Nº Processo': gv('Nº Processo'), 'Estabelecimento': gv('Estabelecimento'),
                'Código APA': gv('Código APA'),
                'Data de Emissão': gv('Data de Início'), 'Data de Entrada em Vigor': gv('Data de Início'),
                'Data de Validade': gv('Data de Validade'), 'Validade (ISO)': val_iso,
                'Em Vigor/Caducada': vig, 'Entidade Licenciadora': gv('Entidade Licenciadora'),
                'Massa de Água': gv('Massa de Água'),
                'Classificação da Massa de Água': gv('Classificação da Massa de Água'),
                'Longitude': gv('Longitude (ETAR)') or gv('Longitude (descarga)'),
                'Latitude': gv('Latitude (ETAR)') or gv('Latitude (descarga)'),
                'Nível de tratamento': gv('Nível de tratamento'),
                'Ano de arranque': gv('Ano de arranque'),
                'Caudal máximo de descarga': gv('Caudal Máximo descarga'),
                'ARH': arh, 'REGIÃO': reg,
            })
            lic_rows.append(row)
            for c in mcond.get(lid, []):
                p = '' if pd.isna(c.get('Parâmetro')) else str(c.get('Parâmetro'))
                if not p:
                    continue
                pk = paramkey(p)
                vle_v = '' if pd.isna(c.get('VLE')) else str(c.get('VLE'))
                vmin, vmax = vle_bounds(vle_v)
                r = {k: '' for k in cond_hdr}
                def cv(col):
                    v = c.get(col, '')
                    return '' if pd.isna(v) else str(v)
                r.update({
                    'Nº TUA': lid, 'Parâmetro': p, 'ParamKey': pk, 'Chave': f'{lid}|{pk}',
                    'VLE (% mín. redução)': cv('VLE (% mín. remoção)'), 'VLE': vle_v,
                    'VLE mín': vmin, 'VLE máx': vmax,
                    'Legislação aplicável': cv('Legislação aplicável'),
                    'Avaliação da Conformidade Legal': cv('Avaliação da Conformidade Legal'),
                    'Frequência de amostragem': cv('Frequência de amostragem'),
                    'Tipo de amostragem': cv('Tipo de amostragem'),
                })
                flags, estado = lurh_criteria(cv('Avaliação da Conformidade Legal'))
                r.update(flags)
                r['Estado do critério'] = estado
                cond_rows.append(r)

    # --- fallback: extracted JSONs cover licences missing from the master ---
    lurh_files = (sorted(glob.glob(os.path.join(LURH_DIR, 'data', 'extracted', '*_extracted.json')))
                  + sorted(glob.glob(os.path.join(LURH_DIR, '*_extracted.json'))))
    for f in lurh_files:
        d = json.load(open(f, encoding='utf-8'))
        g = d.get('dados_gerais', {})
        lid = g.get('Nº Licença', '')
        if not lid or lid in seen:  # skip duplicates/old baselines of same licence
            continue
        seen.add(lid)
        rej = d.get('rejeicao', {}) or {}
        val_iso = iso(g.get('Data de Validade'))
        try:
            vig = 'Em Vigor' if datetime.strptime(val_iso, '%Y-%m-%d').date() >= today else 'Caducada'
        except ValueError:
            vig = ''
        arh, reg = '', ''
        mrh = re.search(r'\.(RH\w+)$', lid)
        if mrh and mrh.group(1) in ARH_BY_RH:
            arh, reg = ARH_BY_RH[mrh.group(1)]
        row = {k: '' for k in lic_hdr}
        row.update({
            'file_name': os.path.basename(f).replace('_extracted.json', '.pdf'),
            'TUA/LURH': 'LURH',
            'Nº TUA': lid,
            'Nº Processo': g.get('Nº Processo', ''),
            'Estabelecimento': rej.get('Designação da rejeição', '') or g.get('Requerente', ''),
            'Código APA': g.get('Código APA', ''),
            'Data de Emissão': g.get('Data de Início', ''),
            'Data de Entrada em Vigor': g.get('Data de Início', ''),
            'Data de Validade': g.get('Data de Validade', ''),
            'Validade (ISO)': val_iso,
            'Em Vigor/Caducada': vig,
            'Município': g.get('Concelho', ''),
            'ARH': arh, 'REGIÃO': reg,
            'Longitude': rej.get('Longitude', ''),
            'Latitude': rej.get('Latitude', ''),
            'Massa de Água': rej.get('Massa de água', ''),
        })
        lic_rows.append(row)
        for c in d.get('condicoes_descarga', []) or []:
            p = c.get('Parâmetro', '')
            if not p: continue
            pk = paramkey(p)
            vmin, vmax = vle_bounds(c.get('VLE'))
            r = {k: '' for k in cond_hdr}
            r.update({
                'Nº TUA': lid, 'Parâmetro': p, 'ParamKey': pk,
                'Chave': f'{lid}|{pk}',
                'VLE (% mín. redução)': c.get('VLE (% mín. remoção)', ''),
                'VLE': c.get('VLE', ''), 'VLE mín': vmin, 'VLE máx': vmax,
                'Legislação aplicável': c.get('Legislação aplicável (texto)') or c.get('Legislação aplicável', ''),
                'Avaliação da Conformidade Legal': c.get('Avaliação da conformidade (texto)', ''),
            })
            flags, estado = lurh_criteria(c.get('Avaliação da conformidade (texto)'))
            r.update(flags)
            r['Estado do critério'] = estado
            cond_rows.append(r)

    def dump(stem, hdr, rows):
        # The dashboard reads .xlsx (one sheet per file, named after the file,
        # all cells as text so Power Query does the date/number coercion).
        #
        # Lock-robust write: build into a temp file, then os.replace() onto the
        # target (atomic — a reader never sees a half-written file). If the target
        # is locked (dashboard open / OneDrive syncing), retry a few times, then
        # fall back to '<stem>.PENDING.xlsx' and warn — never lose the data.
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active; ws.title = stem
        ws.append(hdr)
        for r in rows:
            ws.append([('' if r.get(h) in (None,) else str(r.get(h, ''))) for h in hdr])
        for row in ws.iter_rows():
            for cell in row:
                cell.number_format = '@'
        target = os.path.join(OUT, f'{stem}.xlsx')
        tmp = os.path.join(OUT, f'.{stem}.tmp.xlsx')
        wb.save(tmp)
        import time as _t
        for attempt in range(5):
            try:
                os.replace(tmp, target)
                return
            except PermissionError:
                _t.sleep(1.5 * (attempt + 1))          # transient OneDrive / open-file lock
        pending = os.path.join(OUT, f'{stem}.PENDING.xlsx')
        os.replace(tmp, pending)
        print(f'! {stem}.xlsx locked — wrote {stem}.PENDING.xlsx instead. '
              f'Close Power BI / Excel and re-run make_powerbi_all.py.', file=sys.stderr)

    normalise_numbers(lic_rows)
    # Join key to the ETAR asset registry (Centro Operacional, CODMAXIMO, ...).
    # Built from the licence's own installation name; blank when the name has no
    # counterpart in TabelaCentroide_ETAR, so those licences simply carry no
    # registry attributes instead of being forced onto a wrong asset.
    from etar_key import chave_etar
    centroide_keys = centroide_rows_and_keys()[1]
    for r in lic_rows:
        k = chave_etar(r.get('Estabelecimento') or r.get('Designação') or '')
        r['Chave_ETAR'] = k if k in centroide_keys else ''
    if 'Chave_ETAR' not in lic_hdr:
        lic_hdr = lic_hdr + ['Chave_ETAR']
    dump('Licenses', lic_hdr, lic_rows)

    # The asset registry itself, as a dimension: one row per ETAR, keyed the same
    # way, carrying CODMAXIMO — which is also the key into WWTP_History.
    cen_hdr, _, cen_rows = centroide_rows_and_keys(full=True)
    dump('TabelaCentroide_ETAR', cen_hdr, cen_rows)
    sem_centro = sum(1 for r in lic_rows if not r['Chave_ETAR'])
    print(f'Chave_ETAR: {len(lic_rows) - sem_centro}/{len(lic_rows)} licences matched '
          f'to TabelaCentroide_ETAR ({sem_centro} sem correspondência)')

    # The dashboard reads these files as-is: the parameter is already split into
    # Parametro (abbreviation) / Unidade / Periodo, and limits are normalised, so
    # Power Query only reads and types. See extraction/param_names.py.
    split_param_columns(cond_rows)
    cond_hdr = with_param_columns(cond_hdr)
    cond_hdr = [c for c in cond_hdr if c not in DROP_FROM_CONDITIONS]
    for r in cond_rows:
        r['VLE'] = clean_vle(r.get('VLE'))
        r['VLE máx'] = clean_vle(r.get('VLE máx'))
    dump('Conditions', order_columns('Conditions', cond_hdr), cond_rows)

    # pass-through tables that only exist for TUA (source stays CSV in TUA/)
    for stem in ('Autocontrolo', 'Legislacao', 'Avaliacao'):
        rows = read_csv(f'{stem}.csv')
        hdr = list(rows[0].keys())
        if stem == 'Autocontrolo':
            split_param_columns(rows)
            hdr = with_param_columns(hdr)
        if stem == 'Avaliacao':
            number_avaliacao(rows)
            hdr = hdr + ['Ordem']
        normalise_numbers(rows)
        dump(stem, order_columns(stem, hdr), rows)
    n_lurh = sum(1 for r in lic_rows if r['TUA/LURH'] == 'LURH')
    print(f'powerbi_xlsx_all: {len(lic_rows)} licences ({n_lurh} LURH), {len(cond_rows)} conditions')

if __name__ == '__main__':
    main()
