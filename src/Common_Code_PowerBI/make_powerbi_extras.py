# -*- coding: utf-8 -*-
"""Build the four DERIVED Power BI tables (run AFTER make_powerbi_all.py).

make_powerbi_all.py produces the base tables straight from the extractions:
    Licenses, Conditions, Autocontrolo, Legislacao, Avaliacao
This script produces the four tables that are DERIVED from those (plus the
criteria model and the reference workbooks):

    Licencas_Criterios.xlsx   licence -> criteria bridge (IdCriterio only; the
                              flags themselves live in Catalogo_Criterios)
    Condicoes_Prioritarias.xlsx  'Outras Condicoes (3.21)' clauses that contain
                              one of the priority keywords
    Renovacao_Proposta.xlsx   renewal-request fields, long format
                              (Secção / Campo / Valor), one block per licence
                              whose validity ends within RENEWAL_WINDOW_MONTHS

Every run recomputes the renewal window from TODAY, so licences enter and
leave the renewal list on their own.

Run:  python extraction/make_powerbi_extras.py
It is also called automatically by src/Extraction_Code_TUA/pipeline.py after each batch.
"""
import datetime
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)          # param_names lives next to this file
SRC = os.path.dirname(HERE)           # …/src  (shared code lives in src/Common_Code_PowerBI)
ROOT = os.path.dirname(SRC)           # project root

from param_names import split_parameter, clean_vle   # noqa: E402  (needs sys.path above)
OUT = os.getenv('EPAL_POWERBI_OUT', os.path.join(ROOT, 'data', 'powerbi'))

CRITERIA_MODEL = os.path.join(ROOT, 'data', 'powerbi', 'criteria_table_TUA_v2.xlsx')
CRITERIA_MODEL_LURH = os.path.join(ROOT, 'data', 'powerbi', 'criteria_table_LURH.xlsx')
KEYWORDS_XLSX = os.path.join(ROOT, 'data', 'powerbi', 'TUA and LURH keywords.xlsx')
KEYWORDS_SHEET = 'TUA_LURH - Outras condições'
CENTROIDE_XLSX = os.path.join(ROOT, 'data', 'powerbi', 'TabelaCentroide_ETAR.xlsx')
MASTER_LURH = os.path.join(SRC, 'Extraction_Code_LURH', 'master_lurh.xlsx')
EXTRACTED_DIR = os.path.join(SRC, 'extraction_JSON_TUA')
OUTRAS_SHEET = 'Outras Condicoes (3.21)'

RENEWAL_WINDOW_MONTHS = 8          # flag licences expiring inside this window
MISSING = 'Em falta'               # value shown when no source provides the field

FLAGS = ['Gama de valores (intervalo)', '≤ 100% VLE (dobro)', '≤ 150% VLE',
         '≤ uma ordem de grandeza do VLE', 'média mensal ≤ VLE',
         'média anual ≤ VLE', 'Quadro III DL 152/97 (borlas)']


# --------------------------------------------------------------------------- #
#  helpers
# --------------------------------------------------------------------------- #
def _rows(path, sheet=None):
    """Read a sheet as a list of dicts (first row = header). [] if unreadable."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        print(f'! cannot read {os.path.basename(path)}: {exc}', file=sys.stderr)
        return []
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    it = ws.values
    try:
        hdr = [str(c) if c is not None else '' for c in next(it)]
    except StopIteration:
        wb.close()
        return []
    out = [dict(zip(hdr, r)) for r in it]
    wb.close()
    return out


def dump(stem, hdr, rows):
    """Write <OUT>/<stem>.xlsx, one sheet named after the file, all cells text.

    Lock-robust (same contract as make_powerbi_all.dump): build a temp file,
    then os.replace() atomically. If the target is locked (Power BI / Excel
    open, OneDrive syncing) retry, then fall back to '<stem>.PENDING.xlsx'
    and warn — never lose the data.
    """
    import time
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = stem
    ws.append(hdr)
    for r in rows:
        ws.append(['' if r.get(h) is None else str(r.get(h, '')) for h in hdr])
    for row in ws.iter_rows():
        for cell in row:
            cell.number_format = '@'
    # outputs the dashboard does NOT read go to staging, keeping data/powerbi feed-only
    _dest = os.path.join(ROOT, 'data', '_staging', 'outputs') if stem in ('Renovacao_Proposta',) else OUT
    os.makedirs(_dest, exist_ok=True)
    target = os.path.join(_dest, f'{stem}.xlsx')
    tmp = os.path.join(_dest, f'.{stem}.tmp.xlsx')
    wb.save(tmp)
    for attempt in range(5):
        try:
            os.replace(tmp, target)
            print(f'  {stem}.xlsx  ({len(rows)} rows)')
            return
        except PermissionError:
            time.sleep(1.5 * (attempt + 1))
    os.replace(tmp, os.path.join(_dest, f'{stem}.PENDING.xlsx'))
    print(f'! {stem}.xlsx locked — wrote {stem}.PENDING.xlsx instead. '
          f'Close Power BI / Excel and re-run.', file=sys.stderr)


def norm(s):
    """Accent- and punctuation-insensitive key for matching plant names."""
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().lower()
    s = re.sub(r'\betar\b|\blic\b|\bde\b|\bda\b|\bdo\b|\bdas\b|\bdos\b', ' ', s)
    return re.sub(r'[^a-z0-9]+', '', s)


def parse_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
            try:
                return datetime.datetime.strptime(v[:10], fmt).date()
            except ValueError:
                pass
    return None


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    m += 1
    day = min(d.day, [31, 29 if y % 4 == 0 and (y % 100 or y % 400 == 0) else 28,
                      31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return datetime.date(y, m, day)


def base_param(p):
    """'Azoto total (período de estiagem) (mg/L N)' -> 'Azoto total'."""
    p = re.sub(r'\(per[íi]odo de estiagem\)', '', str(p or ''), flags=re.I)
    return re.sub(r'\s+', ' ', p.split('(')[0]).strip()


def is_estiagem(p):
    return 'estiagem' in str(p or '').lower()


def text_or_missing(v):
    if v is None:
        return MISSING
    s = str(v).strip()
    return s if s and s not in ('0', '(Vide)', 'None') else MISSING


def clean(v):
    """Centroide/master value, or None when it carries no information."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s and s not in ('0', 'None', '-- Não conhecido --') else None


# --------------------------------------------------------------------------- #
#  1. Licencas_Criterios — licence -> criteria bridge (flags denormalised)
# --------------------------------------------------------------------------- #
def build_licencas_criterios():
    # TUA + LURH criteria bridges (disjoint IdCriterio prefixes: TUA_* / LURH_*)
    bridge = _rows(CRITERIA_MODEL, 'Licencas_Criterios') + _rows(CRITERIA_MODEL_LURH, 'Licencas_Criterios')
    if not bridge:
        print('! criteria model unavailable — Licencas_Criterios skipped', file=sys.stderr)
        return []
    rows = []
    for b in bridge:
        lic = str(b.get('IdLicenca') or '').strip()
        if not lic:
            continue
        cid = str(b.get('IdCriterio') or '').strip()
        raw = b.get('Parâmetro (documento)') or ''
        short, unit, period = split_parameter(raw)
        rows.append({'IdLicenca': lic,
                     'Regime': b.get('Regime') or '',
                     'Parametro': short,
                     'Unidade': unit,
                     'Periodo': period,
                     'VLE': clean_vle(b.get('VLE')),
                     'IdCriterio': cid,
                     'Parâmetro (documento)': raw})
    hdr = ['IdLicenca', 'Regime', 'Parametro', 'Unidade', 'Periodo', 'VLE',
           'IdCriterio', 'Parâmetro (documento)']
    dump('Licencas_Criterios', hdr, rows)
    return rows


# --------------------------------------------------------------------------- #
#  3. Condicoes_Prioritarias — 'Outras Condições' clauses matching a keyword
# --------------------------------------------------------------------------- #
def load_keywords():
    kws = []
    for r in _rows(KEYWORDS_XLSX, KEYWORDS_SHEET):
        first = list(r.values())[0] if r else None
        k = str(first or '').strip()
        if k and k.lower() != 'keyword(s)':
            kws.append(k)
    return sorted(set(kws))


def build_condicoes_prioritarias():
    import glob
    import openpyxl
    keywords = load_keywords()
    if not keywords:
        print('! no keywords found — Condicoes_Prioritarias skipped', file=sys.stderr)
        return
    normalised = [(k, norm_text(k)) for k in keywords]
    rows, seen = [], set()
    for path in sorted(glob.glob(os.path.join(EXTRACTED_DIR, '*_extracted.xlsx'))):
        m = re.search(r'(TUA\d+)', os.path.basename(path))
        tua = m.group(1) if m else None
        if not tua:
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception:
            continue
        if OUTRAS_SHEET in wb.sheetnames:
            for r in wb[OUTRAS_SHEET].values:
                if not r or len(r) < 3:
                    continue
                code, clause = r[0], r[2]
                if code in (None, '') or clause in (None, ''):
                    continue
                code = str(code).strip()
                if code in ('Código', '3.21 — Outras Condições'):
                    continue
                hay = norm_text(clause)
                hits = [k for k, kn in normalised if kn in hay]
                if hits and (tua, code) not in seen:
                    seen.add((tua, code))
                    rows.append({'Nº TUA': tua, 'Código': code,
                                 'Palavra-chave': ', '.join(hits),
                                 'Condição': str(clause).strip()})
        wb.close()

    # --- LURH: the same keyword scan, over the 'Outras' clauses already
    # extracted into master_lurh.xlsx's CondicoesTexto sheet (extract_lurh.py's
    # condicoes_texto_rows). LURH clauses have no letter code — the clause's
    # own ordinal ('1ª', '2ª', ...) stands in for 'Código'. ---
    if os.path.isfile(MASTER_LURH):
        import pandas as pd
        try:
            ct = pd.read_excel(MASTER_LURH, sheet_name='CondicoesTexto')
        except Exception:
            ct = None
        if ct is not None and not ct.empty:
            for _, r in ct[ct['Secção'] == 'Outras'].iterrows():
                lic = str(r.get('Nº Licença', '') or '').strip()
                clause = r.get('Condição', '')
                if not lic or clause in (None, ''):
                    continue
                num = r.get('Nº')
                code = f'{int(num)}ª' if num not in (None, '') else ''
                hay = norm_text(clause)
                hits = [k for k, kn in normalised if kn in hay]
                if hits and (lic, code) not in seen:
                    seen.add((lic, code))
                    rows.append({'Nº TUA': lic, 'Código': code,
                                 'Palavra-chave': ', '.join(hits),
                                 'Condição': str(clause).strip()})

    dump('Condicoes_Prioritarias',
         ['Nº TUA', 'Código', 'Palavra-chave', 'Condição'], rows)


def norm_text(s):
    """Accent-insensitive, whitespace-collapsed text (keeps words, for search)."""
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', s).strip()


# --------------------------------------------------------------------------- #
#  4. Renovacao_Proposta — renewal request fields, long format
# --------------------------------------------------------------------------- #
def build_renovacao():
    today = datetime.date.today()
    limit = add_months(today, RENEWAL_WINDOW_MONTHS)

    licences = _rows(os.path.join(OUT, 'Licenses.xlsx'))
    flagged = []
    for r in licences:
        d = parse_date(r.get('Validade (ISO)')) or parse_date(r.get('Data de Validade'))
        if d and today <= d <= limit:
            r['_validade'] = d
            flagged.append(r)
    flagged.sort(key=lambda r: (r['_validade'], str(r.get('Estabelecimento') or '')))

    # reference workbooks: LURH master (discharge point) and centroide (location)
    lurh = {}
    for r in _rows(MASTER_LURH, 'Licenses'):
        key = str(r.get('Nº Licença') or '').strip()
        if key:
            lurh[key] = r
    centroide = {}
    for r in _rows(CENTROIDE_XLSX):
        for col in ('DESIGNACAO', 'DESIGNACAO_GNA', 'SUBSISTEMA', 'RECINTO'):
            key = norm(r.get(col))
            if key:
                centroide.setdefault(key, r)

    def find_centroide(name):
        key = norm(name)
        if key in centroide:
            return centroide[key]
        if len(key) >= 5:                      # tolerate name variants
            for ck, rec in centroide.items():
                if len(ck) >= 5 and (key in ck or ck in key):
                    return rec
        return None

    S1 = 'EXP8.3.5 · Caracterização Geral'
    S2 = 'EXP8.3.7 · Rejeição de águas residuais'
    S3 = 'EXP8.3.8 · Afluente Bruto'
    S4 = 'EXP8.3.11 · Origem'
    S5 = 'Outros'
    S6 = 'População equivalente (últimos 5 anos)'

    def fields(r):
        master = lurh.get(str(r.get('Nº TUA') or '').strip(), {})
        cen = find_centroide(r.get('Estabelecimento'))

        def from_master(*keys):
            for k in keys:
                if master.get(k) not in (None, ''):
                    return str(master[k]).strip()
            return None

        def from_centroide(key):
            return clean(cen.get(key)) if cen else None

        freguesia = from_centroide('FREGUESIA') or MISSING
        concelho = text_or_missing(r.get('Município'))
        if concelho == MISSING:
            concelho = from_centroide('CONCELHO') or MISSING
        ano = text_or_missing(r.get('Ano de arranque'))
        if ano == MISSING:
            ano = from_centroide('ANO') or MISSING
        lon = text_or_missing(r.get('Longitude'))
        if lon == MISSING:
            lon = from_centroide('LONGITUDE_WGS84') or MISSING
        lat = text_or_missing(r.get('Latitude'))
        if lat == MISSING:
            lat = from_centroide('LATITUDE_WGS84') or MISSING

        return [
            (S1, 'Morada', MISSING),
            (S1, 'Início de exploração', from_centroide('ANO') or MISSING),
            (S1, 'Longitude (ETAR)', lon),
            (S1, 'Latitude (ETAR)', lat),
            (S1, 'Designação da instalação',
             text_or_missing(r.get('Designação') or r.get('Estabelecimento'))),
            (S1, 'Distrito', MISSING),
            (S1, 'Concelho', concelho),
            (S1, 'Freguesia', freguesia),
            (S1, 'Ano de arranque', ano),
            (S1, 'População servida à data (ep)',
             text_or_missing(r.get('População servida à data'))),
            (S1, 'Ano horizonte de projeto',
             text_or_missing(r.get('Ano horizonte de projeto'))),
            (S1, 'População servida no horizonte (ep)',
             text_or_missing(r.get('População servida no horizonte'))),
            (S1, 'Nível de tratamento', text_or_missing(r.get('Nível de tratamento'))),
            (S1, 'Esquema de tratamento', text_or_missing(r.get('Esquema de tratamento'))),
            (S1, 'Área (ha)', MISSING),
            (S1, 'Caudal máximo de descarga (m³/dia)',
             text_or_missing(r.get('Caudal máximo de descarga'))),

            (S2, 'Projeto financiado', MISSING),
            (S2, 'Qual (financiamento)', MISSING),
            (S2, 'Comentários (financiamento)', MISSING),
            (S2, 'Designação do ponto de rejeição', from_master('Sistema de Descarga') or MISSING),
            (S2, 'Meio recetor', from_master('Meio Recetor') or MISSING),
            (S2, 'Denominação do meio recetor',
             from_master('Denominação do meio recetor') or text_or_missing(r.get('Massa de Água'))),
            (S2, 'Margem', from_master('Margem') or MISSING),
            (S2, 'Sistema de descarga', from_master('Sistema de Descarga') or MISSING),
            (S2, 'Volume anual descarregado (m³)', MISSING),
            (S2, 'Distrito (rejeição)', MISSING),
            (S2, 'Concelho (rejeição)', concelho),
            (S2, 'Freguesia (rejeição)', freguesia),
            (S2, 'Longitude (rejeição)', from_master('Longitude (descarga)') or MISSING),
            (S2, 'Latitude (rejeição)', from_master('Latitude (descarga)') or MISSING),

            (S3, 'Volume médio mensal (m³)', MISSING),
            (S3, 'CBO5 (mg/L O2)', MISSING),
            (S3, 'CQO (mg/L O2)', MISSING),
            (S3, 'N (mg/L N)', MISSING),
            (S3, 'P (mg/L P)', MISSING),

            (S4, 'Tipo', MISSING),
            (S4, 'Origens', MISSING),
            (S4, 'Instalação de tratamento', text_or_missing(r.get('Estabelecimento'))),

            (S5, 'Centro de Custo da Instalação',
             from_centroide('Centro_Operacional') or MISSING),

            (S6, 'População equivalente 2020', MISSING),
            (S6, 'População equivalente 2021', MISSING),
            (S6, 'População equivalente 2022', MISSING),
            (S6, 'População equivalente 2023', MISSING),
            (S6, 'População equivalente 2024', MISSING),
            (S6, 'Média 2020-2024', MISSING),
        ]

    rows = []
    for r in flagged:
        dias = (r['_validade'] - today).days
        for i, (section, campo, valor) in enumerate(fields(r)):
            rows.append({'Nº TUA': str(r.get('Nº TUA') or ''),
                         'ETAR': str(r.get('Estabelecimento') or ''),
                         'Regime': str(r.get('TUA/LURH') or ''),
                         'Validade': r['_validade'].strftime('%d/%m/%Y'),
                         'Dias p/ expirar': dias,
                         'Secção': section,
                         'Ordem': i,
                         'Campo': campo,
                         'Valor': valor})
    hdr = ['Nº TUA', 'ETAR', 'Regime', 'Validade', 'Dias p/ expirar',
           'Ordem', 'Secção', 'Campo', 'Valor']
    dump('Renovacao_Proposta', hdr, rows)
    missing = sum(1 for r in rows if r['Valor'] == MISSING)
    print(f'  renewal window: {len(flagged)} licences expiring before '
          f'{limit:%d/%m/%Y} — {missing} fields still to fill')


# --------------------------------------------------------------------------- #
def main():
    if not os.path.isdir(OUT):
        print(f'! output folder not found: {OUT}', file=sys.stderr)
        return 1
    print('powerbi_extras: building derived tables')
    build_licencas_criterios()
    build_condicoes_prioritarias()
    build_renovacao()
    print('powerbi_extras: done')
    return 0


if __name__ == '__main__':
    sys.exit(main())
