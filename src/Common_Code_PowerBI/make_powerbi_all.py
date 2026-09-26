# -*- coding: utf-8 -*-
"""Build powerbi_csv_all/ — TUA + LURH licences in one set of Power BI CSVs.

Reads the TUA export (data/powerbi/TUA/*.csv) and appends the LURH
licences (src/extraction_JSON_LURH/*_extracted.json) mapped into the same schema, with
the 'TUA/LURH' column distinguishing the two licence types.
Run after either pipeline updates:  python extraction/make_powerbi_all.py
"""
import csv, glob, json, os, re, sys
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)   # param_names/etar_key/codmaximo_match live next to this file
import licenca_id              # noqa: E402  identidade = nome + periodo
SRC = os.path.dirname(HERE)    # …/src  (shared code lives in src/Common_Code_PowerBI)
ROOT = os.path.dirname(SRC)    # project root
TUA_CSV = os.path.join(ROOT, 'data', 'powerbi', 'TUA')
# NOTE: TUA_CSV used to point at data/_staging/TUA — a one-off, unmanaged copy
# of these same 5 files that nothing ever regenerated. It has since vanished
# (nothing in this pipeline recreates it), while extract_tua.py's own
# export_master_csv() has been writing the live, current versions of these
# exact files straight into data/powerbi/TUA all along. Read from there.
LURH_CODE_DIR = os.path.join(SRC, 'Extraction_Code_LURH')   # code (master lives here)
LURH_JSON_DIR = os.path.join(SRC, 'extraction_JSON_LURH')    # per-file extracted JSON
OUT = os.getenv('EPAL_POWERBI_OUT', os.path.join(ROOT, 'data', 'powerbi'))
CRITERIA_XLSX = os.path.join(ROOT, 'data', 'powerbi', 'Comparacao_Criterios_Conformidade_TODAS_v2.xlsx')
MASTER_LURH = os.path.join(LURH_CODE_DIR, 'master_lurh.xlsx')
CENTROIDE_XLSX = os.path.join(ROOT, 'data', 'powerbi', 'TabelaCentroide_ETAR_newDep_sensity.xlsx')
LURH_CSV = os.path.join(ROOT, 'data', '_staging', 'LURH')
# Working outputs the dashboard does NOT read go to staging, keeping data/powerbi feed-only.
STAGING_OUT = os.path.join(ROOT, 'data', '_staging', 'outputs')
NON_FEED_STEMS = {'TabelaCentroide_ETAR'}

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
    'Comunicacoes': ['Nº TUA', 'Código', 'Tipo de informação/Parâmetros',
                      'Formato de reporte', 'Periodicidade de comunicação',
                      'Data de reporte', 'Entidade'],
    'MeioRecetor': ['Nº TUA', 'Local', 'Parâmetro', 'Método analítico',
                     'Frequência', 'Observações'],
}

# Comunicacoes' TUA side is not one of the 5 normalised MASTER_SHEETS staged in
# TUA_CSV — it is extract_tua.py's own raw per-section export (OCom1), which
# lives in data/powerbi/TUA instead. Read it from there.
COMUNICACOES_TUA_CSV = os.path.join(ROOT, 'data', 'powerbi', 'TUA',
                                     'Obrigacoes Comunicacao (OCom1).csv')
TUA_RAW_DIR = os.path.join(ROOT, 'data', 'powerbi', 'TUA')

# Criteria flags belong to Catalogo_Criterios only — a single source of truth.
# Conditions keeps the licence's own values, not the interpretation.
DROP_FROM_CONDITIONS = CRIT_COLS + CRIT_EXTRA + ['ParamKey', 'Chave']


# Um traco, um 'n/a' ou um 'None' NAO sao dados: sao a maneira de o documento
# dizer que o campo esta vazio. Se seguirem para o feed, aparecem no dashboard
# como se fossem um valor -- '-' numa coluna de metodo analitico parece um
# metodo chamado '-'. Celula vazia e mais honesto: o visual mostra nada.
# ATENCAO: '0' NAO entra nesta lista. Nas colunas de criterios o zero e um
# valor legitimo (0/1 = o criterio nao se aplica / aplica-se).
PLACEHOLDERS = {'-', '--', '---', '—', '–', 'n/a', 'N/A', 'na', 'N.A.', 'n.a.',
                'nan', 'NaN', 'none', 'None', 'null', 'NULL', 's/d', 'S/D',
                'sem dados', '(vide)', '(vazio)', '(blank)', '(em branco)',
                '.', '..', '...', '?', '??'}


def _sem_placeholder(v):
    """O valor tal e qual, ou '' quando e so um marcador de 'nada aqui'."""
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s in PLACEHOLDERS else str(v)


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


def nota_do_tipo_amostragem(tipo, nota):
    """The 'Nota tipo de amostragem' lists every marker's meaning, e.g.:
    '(i) com intervalos máximos de 1 hora; (ii) …; (iii) …; (iv) …'.
    A row's 'Tipo de amostragem' carries ONE marker (e.g. 'Composta (i)').
    Return only that marker's phrase (e.g. '(i) com intervalos máximos de 1 hora'),
    stopping at the next marker / sentence end (drops trailing OCR boilerplate).
    Rows without a marker (Pontual, Em contínuo, …) return ''."""
    import re
    tipo, nota = str(tipo or ''), str(nota or '')
    m = re.search(r'\(([ivx]+)\)', tipo)
    if not m or not nota:
        return ''
    mk = m.group(1)
    seg = re.search(r'\(' + mk + r'\)\s*(.*?)(?=\s*;\s*\([ivx]+\)|\.\s|$)', nota)
    return f'({mk}) ' + seg.group(1).strip().rstrip('. ') if seg else ''


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
    seen_num = set()   # numeros ja trazidos do master
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
            # A identidade e NOME + PERIODO: o mesmo numero pode pertencer a
            # duas licencas diferentes (renovacao, ou a mesma ETAR com outra
            # configuracao), e essas TEM de aparecer as duas.
            ident = licenca_id.identidade(
                g.get('Estabelecimento') or g.get('Designação') or g.get('Requerente'),
                g.get('Data de Início'), g.get('Data de Validade'))
            if not lid or ident in seen:
                continue
            seen.add(ident)
            seen_num.add(str(lid).strip())
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
            # 'Nut III – Concelho – Freguesia' looks like "– Freguesia <NUTIII> / <Concelho> / <Freguesia>"
            _nut = gv('Nut III – Concelho – Freguesia')
            municipio = _nut.split('/')[1].strip() if _nut.count('/') >= 2 else ''
            row = {k: '' for k in lic_hdr}
            row.update({
                'file_name': gv('file_name'), 'TUA/LURH': 'LURH', 'Nº TUA': lid,
                licenca_id.COLUNA: ident,
                'Nº Processo': gv('Nº Processo'), 'Estabelecimento': gv('Estabelecimento'),
                # LURH source has no separate 'Designação' — it's the plant name, same as Estabelecimento
                'Designação': gv('Estabelecimento'),
                'Código APA': gv('Código APA'),
                'Data de Emissão': gv('Data de Início'), 'Data de Entrada em Vigor': gv('Data de Início'),
                'Data de Validade': gv('Data de Validade'), 'Validade (ISO)': val_iso,
                'Em Vigor/Caducada': vig, 'Entidade Licenciadora': gv('Entidade Licenciadora'),
                'Massa de Água': gv('Massa de Água'),
                'Meio Recetor': gv('Meio Recetor'),
                'Denominação do meio recetor': gv('Denominação do meio recetor'),
                'Classificação da Massa de Água': gv('Classificação da Massa de Água'),
                'Longitude': gv('Longitude (ETAR)') or gv('Longitude (descarga)'),
                'Latitude': gv('Latitude (ETAR)') or gv('Latitude (descarga)'),
                'Nível de tratamento': gv('Nível de tratamento'),
                'Esquema de tratamento': gv('Tipo de tratamento'),
                'Tratamento implementado': gv('Nível de tratamento'),
                'Ano de arranque': gv('Ano de arranque'),
                'Ano horizonte de projeto': gv('Ano horizonte de projeto'),
                'População servida à data': gv('População servida (e.p.)'),
                'População servida no horizonte': gv('População servida no horizonte'),
                'Caudal máximo de descarga': gv('Caudal Máximo descarga'),
                'Código TURH': lid,
                'Município': municipio,
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
    lurh_files = sorted(glob.glob(os.path.join(LURH_JSON_DIR, '*_extracted.json')))
    for f in lurh_files:
        d = json.load(open(f, encoding='utf-8'))
        g = d.get('dados_gerais', {})
        lid = g.get('Nº Licença', '')
        rej = d.get('rejeicao', {}) or {}
        ident = licenca_id.identidade(
            rej.get('Designação da rejeição') or g.get('Requerente'),
            g.get('Data de Início'), g.get('Data de Validade'))
        if not lid or ident in seen or str(lid).strip() in seen_num:
            continue                       # ja veio do master, ou ja foi visto
        seen.add(ident)
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
            licenca_id.COLUNA: ident,
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
            'Meio Recetor': rej.get('Meio Recetor', ''),
            'Denominação do meio recetor': rej.get('Denominação do meio recetor', ''),
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
            ws.append([_sem_placeholder(r.get(h)) for h in hdr])
        for row in ws.iter_rows():
            for cell in row:
                cell.number_format = '@'
        _dest = STAGING_OUT if stem in NON_FEED_STEMS else OUT
        os.makedirs(_dest, exist_ok=True)
        target = os.path.join(_dest, f'{stem}.xlsx')
        tmp = os.path.join(_dest, f'.{stem}.tmp.xlsx')
        wb.save(tmp)
        import time as _t
        for attempt in range(5):
            try:
                os.replace(tmp, target)
                return
            except PermissionError:
                _t.sleep(1.5 * (attempt + 1))          # transient OneDrive / open-file lock
        pending = os.path.join(_dest, f'{stem}.PENDING.xlsx')
        os.replace(tmp, pending)
        print(f'! {stem}.xlsx locked — wrote {stem}.PENDING.xlsx instead. '
              f'Close Power BI / Excel and re-run make_powerbi_all.py.', file=sys.stderr)

    normalise_numbers(lic_rows)
    # Canonicalise 'Em Vigor/Caducada' casing: extract_tua.py writes 'Em vigor'
    # (lowercase v) while the LURH computation above writes 'Em Vigor' (capital
    # V) — same status, two different strings, which silently splits one
    # filter/slicer value into two in Power BI. 'Caducada' and blanks are
    # already consistent and left untouched.
    for r in lic_rows:
        if (r.get('Em Vigor/Caducada') or '').strip().lower() == 'em vigor':
            r['Em Vigor/Caducada'] = 'Em Vigor'
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
    # 'REGIÃO'/'ARH'/'Município' are computed above (per LURH row, from the RH
    # code) but were never added to the written header — dump() only writes
    # columns present in lic_hdr, so these silently vanished from Licenses.xlsx
    # even though the Power Query model (Preenchidos step) expects them.
    for col in ('Meio Recetor', 'Denominação do meio recetor', 'REGIÃO', 'ARH', 'Município'):
        if col not in lic_hdr:
            lic_hdr = lic_hdr + [col]
    # CODMAXIMO — the asset key into the ETAR registry. Matched by the licence's
    # installation name (Stº->Santo, articles/ETAR folded), with the nearest
    # registry asset within 2 km as fallback; name wins so a wrong coordinate
    # cannot pull a licence onto the wrong plant. Source of truth is
    # TabelaCentroide_ETAR_newDep_sensity.xlsx (see codmaximo_match.py).
    from codmaximo_match import load_registry, match_codmaximo
    _reg_cod = load_registry()
    _cod_hits = 0
    for r in lic_rows:
        code, _m = match_codmaximo(r.get('Designação'), r.get('Latitude'), r.get('Longitude'), _reg_cod)
        r['CODMAXIMO'] = code
        _cod_hits += 1 if code else 0
    if 'CODMAXIMO' not in lic_hdr:
        lic_hdr = lic_hdr + ['CODMAXIMO']
    print(f'CODMAXIMO: {_cod_hits}/{len(lic_rows)} licences matched to the ETAR registry')

    # 'Vigência da Licença' — when an ETAR (grouped by CODMAXIMO) has more than
    # one licence on file (e.g. an old one plus its renewal), tag the one with
    # the most recent 'Data de Entrada em Vigor' as 'Atual' and the rest as
    # 'Histórico'. This is deliberately separate from 'Em Vigor/Caducada':
    # that column says whether THIS licence's own dates have lapsed, not
    # whether it is the newest one for its plant — an ETAR's latest licence
    # could in principle already be expired with no renewal processed yet,
    # and it should still read as 'Atual' for that plant.
    from collections import defaultdict
    _by_etar = defaultdict(list)
    for _i, r in enumerate(lic_rows):
        key = r.get('CODMAXIMO') or f'__sem_codmaximo_{_i}'   # no registry match -> stands alone
        _by_etar[key].append(r)

    def _entry_date(row):
        # Missing 'Data de Entrada em Vigor' sorts last (never beats a dated
        # sibling for 'Atual'); with only one licence in the group it still
        # correctly ends up 'Atual' regardless.
        return iso(row.get('Data de Entrada em Vigor')) or '0000-00-00'

    for _rows in _by_etar.values():
        _rows.sort(key=_entry_date, reverse=True)   # ties: stable, keeps original order
        for _i, row in enumerate(_rows):
            row['Vigência da Licença'] = 'Atual' if _i == 0 else 'Histórico'
    if 'Vigência da Licença' not in lic_hdr:
        lic_hdr = lic_hdr + ['Vigência da Licença']
    _n_hist = sum(1 for r in lic_rows if r.get('Vigência da Licença') == 'Histórico')
    print(f"Vigência da Licença: {_n_hist} licence(s) marked Histórico (superseded within the same ETAR)")

    # Datas: um marcador de posição do documento ('-', 'n/a', ...) não é uma
    # data. Se seguir para o feed, o Power BI falha a conversão da coluna
    # inteira com "Le type de la valeur ne correspond pas à celui de la
    # colonne" e cria tabelas 'Erreurs dans Licences' — foi o que aconteceu
    # com o TUA20231012002973 (ETAR Vera Cruz), que traz '-' na validade.
    # Uma célula vazia é lida como data em branco e não parte nada.
    _DATAS = ('Data de Emissão', 'Data de Entrada em Vigor', 'Data de Validade')
    _VAZIOS = {'-', '--', '---', '—', 'n/a', 'N/A', 'n.a.', 's/d', '.'}
    _limpas = 0
    for _r in lic_rows:
        for _c in _DATAS:
            _v = str(_r.get(_c, '')).strip()
            if _v and (_v in _VAZIOS or not re.match(r'^\d{2}[-/]\d{2}[-/]\d{4}$', _v)):
                _r[_c] = ''
                _limpas += 1
    if _limpas:
        print(f'Datas: {_limpas} valor(es) que não são datas foram postos a vazio '
              f'(evita o erro de conversão no Power BI).')

    # Ligação clicável para o PDF original, para o botão "Abrir licença" do
    # dashboard. Vazia quando o ficheiro não está no arquivo — o botão
    # desliga-se sozinho nesse caso, em vez de abrir um endereço partido.
    import pdf_link
    _com, _sem = pdf_link.acrescentar_coluna(lic_rows, os.path.join(ROOT, 'data', 'pdfs'))
    if pdf_link.COLUNA not in lic_hdr:
        lic_hdr = lic_hdr + [pdf_link.COLUNA]
    _base = pdf_link.base_url()
    print(f"{pdf_link.COLUNA}: {_com}/{_com + _sem} licences link to their PDF"
          + (f' (base: {_base})' if _base
             else ' (file:/// local — só abre no Power BI Desktop;'
                  ' definir EPAL_PDF_BASE_URL depois da mudança para o SharePoint)'))

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
    # Diploma(s) citados na Legislação aplicável (todos, ex.: "Lei 58/2005; DL 152/97").
    # Espelha o lado da Avaliação; parser em diplomas.py. Coluna nova → order_columns
    # mantém-na no fim; a Ficha da Conformidade Legal mostra-a a seguir a "Unidade".
    try:
        import diplomas as _dip
        for r in cond_rows:
            _t = str(r.get('Legislação aplicável') or '').strip()
            r['Diploma (Legislação)'] = ('; '.join(d['diploma'] for d in _dip.extract(_t))
                                         if _t else '')
        if 'Diploma (Legislação)' not in cond_hdr:
            cond_hdr.append('Diploma (Legislação)')
    except Exception as _exc:
        print(f'  ! coluna "Diploma (Legislação)" ignorada: {_exc}')
    dump('Conditions', order_columns('Conditions', cond_hdr), cond_rows)

    # these tables exist for BOTH TUA and LURH — combine them.
    # LURH rows come from data/powerbi/LURH/<stem>.csv (exported above from
    # master_lurh.xlsx); their licence key is 'Nº Licença' -> map to 'Nº TUA'.
    for stem in ('Autocontrolo', 'Legislacao', 'Avaliacao', 'Comunicacoes', 'MeioRecetor'):
        if stem == 'Comunicacoes':
            if os.path.isfile(COMUNICACOES_TUA_CSV):
                with open(COMUNICACOES_TUA_CSV, encoding='utf-8-sig', newline='') as _f:
                    rows = list(csv.DictReader(_f))
            else:
                rows = []
            hdr = list(rows[0].keys()) if rows else list(COLUMN_ORDER['Comunicacoes'])
        elif stem == 'MeioRecetor':
            # TUA licences have no equivalent 'monitoring of the receiving
            # water body' section — this table is LURH-only, built entirely
            # from LURH_CSV below.
            rows, hdr = [], list(COLUMN_ORDER['MeioRecetor'])
        else:
            rows = read_csv(f'{stem}.csv')
            hdr = list(rows[0].keys())
        lurh_p = os.path.join(LURH_CSV, f'{stem}.csv')
        if os.path.isfile(lurh_p):
            with open(lurh_p, encoding='utf-8-sig', newline='') as _f:
                lrows = list(csv.DictReader(_f))
            for lr in lrows:
                if 'Nº Licença' in lr:
                    lr['Nº TUA'] = lr.pop('Nº Licença')
                rows.append(lr)
            for c in (list(lrows[0].keys()) if lrows else []):
                if c not in hdr:
                    hdr.append(c)
        if stem in ('Autocontrolo', 'MeioRecetor'):
            split_param_columns(rows)
            hdr = with_param_columns(hdr)
        if stem == 'Avaliacao':
            number_avaliacao(rows)
            hdr = hdr + ['Ordem']
        if stem == 'Autocontrolo':
            for r in rows:
                r['Nota do tipo de amostragem'] = nota_do_tipo_amostragem(
                    r.get('Tipo de amostragem', ''), r.get('Nota tipo de amostragem', ''))
            if 'Nota do tipo de amostragem' not in hdr:
                hdr.append('Nota do tipo de amostragem')
        normalise_numbers(rows)
        dump(stem, order_columns(stem, hdr), rows)
    # --- Pesquisa: one unified full-text search index across BOTH regimes —
    # every clause of Condições Gerais/Específicas/Outras Condições plus every
    # Obrigação de Comunicação, tagged with its licence and its section, all
    # in a single 'Texto' column. This is what the dashboard's single search
    # box (a slicer with search enabled on Pesquisa[Texto]) reads — one field
    # to bind, instead of one slicer per source table. ---
    def read_csv_raw(path):
        if not os.path.isfile(path):
            return []
        with open(path, encoding='utf-8-sig', newline='') as _f:
            return list(csv.DictReader(_f))

    pesquisa_rows = []
    TUA_SOURCES = [
        ('Condicoes Gerais (3.19).csv', 'Condições Gerais'),
        ('Condicoes Especificas (3.20).csv', 'Condições Específicas'),
        ('Outras Condicoes (3.21).csv', 'Outras Condições'),
    ]
    for fname, secao in TUA_SOURCES:
        for r in read_csv_raw(os.path.join(TUA_RAW_DIR, fname)):
            texto = (r.get('Condição') or '').strip()
            if texto:
                pesquisa_rows.append({'Nº TUA': r.get('Nº TUA', ''), 'Secção': secao,
                                      'Código': r.get('Código', ''), 'Texto': texto})
    for r in read_csv_raw(COMUNICACOES_TUA_CSV):
        tipo = (r.get('Tipo de informação/Parâmetros') or '').strip()
        periodo = (r.get('Periodicidade de comunicação') or '').strip()
        texto = ' — '.join(t for t in (tipo, periodo) if t)
        if texto:
            pesquisa_rows.append({'Nº TUA': r.get('Nº TUA', ''), 'Secção': 'Obrigação de Comunicação',
                                  'Código': r.get('Código', ''), 'Texto': texto})

    if os.path.isfile(MASTER_LURH):
        import pandas as pd
        ml2 = pd.read_excel(MASTER_LURH, sheet_name=None)
        for _, r in ml2.get('CondicoesTexto', pd.DataFrame()).iterrows():
            texto = str(r.get('Condição', '') or '').strip()
            if texto and texto.lower() != 'nan':
                secao = {'Gerais': 'Condições Gerais', 'Específicas': 'Condições Específicas',
                         'Outras': 'Outras Condições'}.get(str(r.get('Secção', '')), str(r.get('Secção', '')))
                pesquisa_rows.append({'Nº TUA': str(r.get('Nº Licença', '')), 'Secção': secao,
                                      'Código': str(r.get('Nº', '')), 'Texto': texto})
        for _, r in ml2.get('Comunicacoes', pd.DataFrame()).iterrows():
            tipo = str(r.get('Tipo de informação/Parâmetros', '') or '').strip()
            periodo = str(r.get('Periodicidade de comunicação', '') or '').strip()
            tipo = '' if tipo.lower() == 'nan' else tipo
            periodo = '' if periodo.lower() == 'nan' else periodo
            texto = ' — '.join(t for t in (tipo, periodo) if t)
            if texto:
                pesquisa_rows.append({'Nº TUA': str(r.get('Nº Licença', '')),
                                      'Secção': 'Obrigação de Comunicação',
                                      'Código': '', 'Texto': texto})

    dump('Pesquisa', ['Nº TUA', 'Secção', 'Código', 'Texto'], pesquisa_rows)
    print(f'Pesquisa: {len(pesquisa_rows)} searchable snippets')

    # 'Id Versão' e a chave por onde o modelo se relaciona (o numero da licenca
    # deixou de ser unico quando as renovacoes passaram a ter linha propria).
    # As tabelas vem de origens diferentes e nem todas a trazem; este passo
    # fecha a diferenca num sitio so. Ver completar_id_versao.py.
    import completar_id_versao
    print('Id Versão: a completar nas tabelas ligadas a uma licença...')
    completar_id_versao.completar(OUT)

    n_lurh = sum(1 for r in lic_rows if r['TUA/LURH'] == 'LURH')
    print(f'powerbi_xlsx_all: {len(lic_rows)} licences ({n_lurh} LURH), {len(cond_rows)} conditions')

if __name__ == '__main__':
    main()
