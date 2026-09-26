"""Sync new 'Avaliação da Conformidade Legal' texts from the master LURH
workbook into the criteria workbook (AdVT_CriteriosConformidade +
CriteriosConformidade_GLOBAL).

How it works
------------
1. Reads every (Parâmetro, Legislação, Avaliação) combination from the
   master's 'Conditions' sheet.
2. Compares them (accent/case/whitespace-insensitive) with rows already in
   AdVT_CriteriosConformidade.
3. For each NEW combination, classifies the Avaliação text into Critérios 1-7
   with keyword rules, appends a row to AdVT_CriteriosConformidade and (once
   per distinct text) a flag row to CriteriosConformidade_GLOBAL.
4. New rows carry the marker 'AUTO — rever' so a human validates them.
5. A timestamped .bak_ copy of the criteria workbook is made before saving.

Run standalone:  python sync_criterios.py
Or it is called automatically by pipeline.py after each master update.
"""
import os
import re
import shutil
import unicodedata
from datetime import datetime

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
MASTER = os.getenv('EPAL_LURH_MASTER', os.path.join(HERE, 'master_lurh.xlsx'))
CRITERIA = os.getenv('EPAL_CRITERIA', os.path.join(
    HERE, 'CriteriosConformidade_geral_20250926_funcionalidades_critérios.xlsx'))

ADVT_SHEET = 'AdVT_CriteriosConformidade'
GLOBAL_SHEET = 'CriteriosConformidade_GLOBAL'


# ------------------------------------------------------------------ helpers
def norm(t):
    """Lowercase, strip accents, collapse whitespace -> comparison key."""
    if not t:
        return ''
    t = unicodedata.normalize('NFKD', str(t))
    t = ''.join(c for c in t if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', t).strip().lower()


def classify(txt):
    """Keyword rules -> which Critérios the Avaliação text implies."""
    t = norm(txt)
    return {
        'C1_media_mensal': 'media mensal' in t,
        'C5_media_anual':  'media anual' in t,
        'C2_dobro':        'dobro do vle' in t,
        'C3_borlas':       ('152/97' in t and (
                             'amostras nao conformes' in t
                             or 'minimo anual de amostras' in t
                             or 'quadro n' in t
                             or 'alinea d' in t)),
        'C4_100':          ('100%' in t or 'mais de 100' in t),
        'C4_150':          '150%' in t,
        'C6_ordem':        'ordem de grandeza' in t,
        'C7_intervalo':    bool(re.search(r'5[.,]0\s*-?\s*10[.,]0', t)),
        'art69_only':      (('artigo 69' in t or 'art. 69' in t
                             or 'artigo n' in t and '69' in t)
                            and '152/97' not in t),
    }


# ------------------------------------------------- read master combinations
def master_triples():
    wb = openpyxl.load_workbook(MASTER, read_only=True)
    ws = wb['Conditions']
    hdr = [c.value for c in next(ws.iter_rows(max_row=1))]
    ix = {h: i for i, h in enumerate(hdr)}
    need = ('Parâmetro', 'Legislação aplicável', 'Avaliação da Conformidade Legal')
    missing = [h for h in need if h not in ix]
    if missing:
        raise RuntimeError(f"Master 'Conditions' sheet is missing columns: {missing}")
    triples = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        par, leg, av = (row[ix[h]] for h in need)
        if par and av:
            triples[(norm(par), norm(leg), norm(av))] = (str(par), str(leg or ''), str(av))
    wb.close()
    return triples


def existing_advt_keys(ws):
    keys = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] and row[2]:
            keys.add((norm(row[0]), norm(row[1]), norm(row[2])))
    return keys



# ---------------------------------------------- formatting-preserving append
from copy import copy
from openpyxl.utils import get_column_letter, range_boundaries


def _style_row(ws, ref_row):
    """Snapshot of each cell style in a reference data row."""
    return {c: ws.cell(row=ref_row, column=c) for c in range(1, ws.max_column + 1)}


def _append_styled(ws, values, styles):
    """Append values and copy font/border/alignment/fill/number format from
    the reference row so the new row matches the original layout."""
    r = ws.max_row + 1
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=r, column=c, value=v)
        ref = styles.get(c)
        if ref is not None and ref.has_style:
            cell.font = copy(ref.font)
            cell.border = copy(ref.border)
            cell.alignment = copy(ref.alignment)
            cell.fill = copy(ref.fill)
            cell.number_format = ref.number_format
    return r


def _extend_filter(ws, last_row):
    """Grow the sheet AutoFilter (and any Excel Table) down to last_row so
    appended rows are visible in the existing filter dropdowns."""
    if ws.auto_filter.ref:
        c1, r1, c2, _ = range_boundaries(ws.auto_filter.ref)
        ws.auto_filter.ref = (f"{get_column_letter(c1)}{r1}:"
                              f"{get_column_letter(c2)}{last_row}")
    for t in getattr(ws, 'tables', {}).values():
        c1, r1, c2, _ = range_boundaries(t.ref)
        t.ref = (f"{get_column_letter(c1)}{r1}:"
                 f"{get_column_letter(c2)}{last_row}")


def _last_data_row(ws, cols=(1, 2, 3)):
    """Last row that actually contains data in the given columns (ignores
    leftover empty-but-formatted rows that inflate ws.max_row)."""
    last = 1
    for r in range(1, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, '') for c in cols):
            last = r
    return last


def _write_styled(ws, row, values, styles):
    """Write values at a specific row (overwriting blanks), copying styles."""
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=c, value=v)
        ref = styles.get(c)
        if ref is not None and ref.has_style:
            cell.font = copy(ref.font)
            cell.border = copy(ref.border)
            cell.alignment = copy(ref.alignment)
            cell.fill = copy(ref.fill)
            cell.number_format = ref.number_format


def _extend_cf(ws, last_row):
    """Grow every conditional-formatting range down to last_row so the
    0/1 colour rules apply to the appended rows."""
    from openpyxl.formatting.formatting import ConditionalFormattingList
    old = list(ws.conditional_formatting)
    ws.conditional_formatting = ConditionalFormattingList()
    for fmt in old:
        c1, r1, c2, _ = range_boundaries(str(fmt.sqref).split()[0])
        new_ref = (f"{get_column_letter(c1)}{r1}:"
                   f"{get_column_letter(c2)}{last_row}")
        for rule in fmt.rules:
            ws.conditional_formatting.add(new_ref, rule)

# ------------------------------------------------------------------- sync
def sync(dry_run=False):
    triples = master_triples()
    wb = openpyxl.load_workbook(CRITERIA)
    advt, glob = wb[ADVT_SHEET], wb[GLOBAL_SHEET]
    seen = existing_advt_keys(advt)
    advt_styles = _style_row(advt, 2)          # row 2 = first original data row
    glob_styles = _style_row(glob, 2)
    advt_next = _last_data_row(advt) + 1       # write right after real data,
    glob_next = _last_data_row(glob) + 1       # not after empty formatted rows
    seen_glob = {norm(r[12]) for r in glob.iter_rows(min_row=2, values_only=True)
                 if r and len(r) > 12 and r[12]}

    added = []
    for key, (par, leg, av) in sorted(triples.items()):
        if key in seen:
            continue
        h = classify(av)
        _write_styled(advt, advt_next, [
            par, leg, av,
            'Média mensal =< VLE' if h['C1_media_mensal'] else None,
            'Valor máximo observado ≤ ao dobro do VLE'
                if (h['C2_dobro'] or h['art69_only']) else None,
            'Borlas' if h['C3_borlas'] else None,
            ('Valor Máximo Observado não superior a 150% VLE' if h['C4_150']
             else 'Valor Máximo Observado não superior a 100% VLE'
                if h['C4_100'] else None),
            'Média anual ≤ VLE' if h['C5_media_anual'] else None,
            'Uma ordem de grandeza' if h['C6_ordem'] else None,
            'Intervalo 5,0-10,0' if h['C7_intervalo'] else None,
            'AdVT — rever'  # marker for human validation,
        ], advt_styles)
        advt_next += 1
        if norm(av) not in seen_glob:
            new_id = f"{re.sub(r'[^A-Za-z0-9]', '', par)[:8]}_{glob_next}"
            _write_styled(glob, glob_next, [
                new_id, 'AdVT', par,
                1, 0,
                1 if h['C7_intervalo'] else 0,
                1 if (h['C4_100'] or h['C2_dobro'] or h['art69_only']) else 0,
                1 if h['C4_150'] else 0,
                1 if h['C6_ordem'] else 0,
                1 if h['C1_media_mensal'] else 0,
                1 if h['C5_media_anual'] else 0,
                1 if h['C3_borlas'] else 0,
                av,
            ], glob_styles)
            glob_next += 1
            seen_glob.add(norm(av))
        seen.add(key)
        added.append((par, av))

    if added and not dry_run:
        _extend_filter(advt, advt_next - 1)
        _extend_filter(glob, glob_next - 1)
        _extend_cf(glob, glob_next - 1)
        shutil.copy2(CRITERIA, CRITERIA + f'.bak_{datetime.now():%Y%m%d_%H%M%S}')
        try:
            wb.save(CRITERIA)
        except PermissionError:
            alt = CRITERIA.replace('.xlsx', f'_PENDING_{datetime.now():%H%M%S}.xlsx')
            wb.save(alt)
            print(f'AVISO: ficheiro de critérios aberto no Excel — gravado em {alt}')
    wb.close()
    return added


if __name__ == '__main__':
    import sys
    dry = '--dry-run' in sys.argv
    added = sync(dry_run=dry)
    tag = ' (simulação, nada gravado)' if dry else ''
    if added:
        print(f'{len(added)} novas combinações adicionadas{tag}:')
        for p, a in added:
            print(f'  - {p}: {a[:70]}...')
    else:
        print('Nenhum critério novo — tudo sincronizado.')