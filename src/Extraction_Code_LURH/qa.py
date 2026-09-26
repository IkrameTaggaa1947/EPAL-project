"""Quality checks for a LURH extraction.

Returns a verdict — OK / FAILED — plus human-readable reasons. The REVIEW verdict
was removed: only a genuine FAILED (Nº Licença missing, or contradicting the
filename) keeps a licence out of the master. Missing sections — no discharge
conditions, no autocontrolo rows, etc. — are informational notes only (logged +
in qa_report.xlsx); a licence with some tables missing is still added, since a
human can fill the gap later and there's no reason to withhold the rest of what
WAS extracted correctly.

The strongest check is free: the SILIAMB filename encodes the ground truth —
    Belmonte_L005347_2021_RH5A_12_07_2021_a_11_07_2026.pdf
             └──── Nº Licença ────┘ └─ início ─┘   └ validade ┘
so we can verify the extracted Nº Licença (L005347.2021.RH5A) and both dates
against the filename with no external data. The regex tolerates the filename
quirks seen in the corpus: a fused 'RH5A02_09_2019' (missing underscore) and
compact dates '27032014_a_27032021' (no separators).
"""
import re
from datetime import datetime

# Nº Licença parts + both dates, with optional/missing separators tolerated
FNAME_RE = re.compile(
    r'L(\d{6})_(\d{4})_(RH\d+A?)_?'
    r'(\d{2})[._-]?(\d{2})[._-]?(\d{4})_a_'
    r'(\d{2})[._-]?(\d{2})[._-]?(\d{4})')
DATE_RE = re.compile(r'^\d{2}-\d{2}-\d{4}$')
VLE_RE = re.compile(r'^\s*\d+([.,]\d+)?(\s*[-a]\s*\d+([.,]\d+)?)?\s*$')

OK, REVIEW, FAILED = 'OK', 'REVIEW', 'FAILED'


def from_filename(name):
    """Ground-truth (Nº Licença, início, validade) parsed from the filename,
    or None if the filename isn't in the expected SILIAMB pattern."""
    m = FNAME_RE.search(name or '')
    if not m:
        return None
    num, year, rh, d1, m1, y1, d2, m2, y2 = m.groups()
    return {'Nº Licença': f'L{num}.{year}.{rh}',
            'Data de Início': f'{d1}-{m1}-{y1}',
            'Data de Validade': f'{d2}-{m2}-{y2}'}


def _num(s):
    try:
        return float(str(s).replace(',', '.'))
    except (TypeError, ValueError):
        return None


def check(d):
    """Return (verdict, reasons). reasons is a list of '[LEVEL] message'."""
    g, t, r = d['dados_gerais'], d['tratamento'], d['rejeicao']
    reasons = []

    def flag(level, msg):
        reasons.append((level, msg))

    # --- hard failures --------------------------------------------------------
    # Only a missing/contradicted licence identity is fatal. A missing TABLE
    # (no conditions, no autocontrolo, ...) is a gap to flag, not a reason to
    # keep the whole licence out of the master — whatever else WAS extracted
    # is still correct and still useful.
    if not g.get('Nº Licença'):
        flag(FAILED, 'Nº Licença missing')

    # --- filename cross-check (ground truth) --------------------------------
    # Only applies when the filename actually IS in the SILIAMB pattern — a
    # PDF renamed by hand (or from another source) just skips the cross-check
    # silently. That's an expected, normal case, not something worth a note
    # on every single drop that doesn't happen to use that naming scheme.
    fn = from_filename(d.get('file_name', ''))
    if fn:
        if g.get('Nº Licença') and g['Nº Licença'] != fn['Nº Licença']:
            flag(FAILED, f"Nº Licença mismatch: extracted {g['Nº Licença']} "
                         f"vs filename {fn['Nº Licença']}")
        for k in ('Data de Início', 'Data de Validade'):
            got, exp = g.get(k), fn[k]
            if not got:
                flag(REVIEW, f"{k} missing (filename says {exp})")
            elif got != exp:
                flag(REVIEW, f"{k} {got} does not match filename {exp}")

    # --- date validity ------------------------------------------------------
    for k in ('Data de Início', 'Data de Validade'):
        v = g.get(k)
        if v and not DATE_RE.match(v):
            flag(REVIEW, f"{k} is not a valid date: {v!r}")

    # --- sections present ---------------------------------------------------
    if not d.get('condicoes_descarga'):
        flag(REVIEW, 'no discharge conditions (VLE table) were extracted')
    if not d.get('autocontrolo', {}).get('rows'):
        flag(REVIEW, 'no self-monitoring (autocontrolo) rows extracted')

    # --- coordinates inside mainland Portugal bounds ------------------------
    for label, sec in (('descarga', r), ('ETAR', t)):
        lon, lat = _num(sec.get('Longitude')), _num(sec.get('Latitude'))
        if lon is not None and not (-10.0 <= lon <= -6.0):
            flag(REVIEW, f'longitude ({label}) out of expected range: {lon}')
        if lat is not None and not (36.5 <= lat <= 42.5):
            flag(REVIEW, f'latitude ({label}) out of expected range: {lat}')

    # --- VLE sanity ---------------------------------------------------------
    for c in d.get('condicoes_descarga', []):
        for field in ('VLE', 'VLE (% mín. remoção)'):
            v = (c.get(field) or '').replace(' a ', '-')
            if v and not VLE_RE.match(v):
                flag(REVIEW, f"unusual {field} for {c.get('Parâmetro', '?')[:24]}: {c.get(field)!r}")

    # --- carry the extractor's own warnings ---------------------------------
    for w in d.get('_warnings', []):
        if 'garbled' in w or 'fragment' in w or 'confirm' in w or 'remoção' in w:
            flag(REVIEW, w)

    # REVIEW step removed: only a genuine FAILED keeps a licence out of the master.
    # All the lower-level checks still run and are returned as informational notes
    # (logged + in qa_report.xlsx), they just no longer divert files to review/.
    verdict = FAILED if any(l == FAILED for l, _ in reasons) else OK
    return verdict, [f'[{l}] {m}' for l, m in reasons]


# ------------------------------------------------------------------ self-check
def _demo():
    good = {'file_name': 'Belmonte_L005347_2021_RH5A_12_07_2021_a_11_07_2026.pdf',
            'dados_gerais': {'Nº Licença': 'L005347.2021.RH5A',
                             'Data de Início': '12-07-2021',
                             'Data de Validade': '11-07-2026'},
            'tratamento': {'Longitude': '-7.35', 'Latitude': '40.36'},
            'rejeicao': {'Longitude': '-7.34', 'Latitude': '40.35'},
            'condicoes_descarga': [{'Parâmetro': 'pH', 'VLE': '6-9'}],
            'autocontrolo': {'rows': [{'x': 1}]}, '_warnings': []}
    v, _ = check(good)
    assert v == OK, v

    # compact-date 2014 filename must parse too
    fn = from_filename('Aldeia_da_Serra_Lic_L003886_2014_RH7_27032014_a_27032021.pdf')
    assert fn == {'Nº Licença': 'L003886.2014.RH7',
                  'Data de Início': '27-03-2014', 'Data de Validade': '27-03-2021'}, fn
    # fused RH+date filename must parse too
    fn = from_filename('Alcorrego__Avis_L014144_2019_RH5A02_09_2019_a_01_09_2024.pdf')
    assert fn == {'Nº Licença': 'L014144.2019.RH5A',
                  'Data de Início': '02-09-2019', 'Data de Validade': '01-09-2024'}, fn

    bad = dict(good, dados_gerais={'Nº Licença': 'L005347.2021.RH5A'})   # dates lost
    v, reasons = check(bad)
    assert v == OK, (v, reasons)                  # REVIEW step removed -> not blocked
    assert any('Validade' in r for r in reasons)  # but the issue is still recorded as a note

    broken = dict(good, condicoes_descarga=[])    # no conditions — a gap, not a failure now
    v, reasons = check(broken)
    assert v == OK, (v, reasons)
    assert any('discharge conditions' in r for r in reasons)

    no_id = dict(good, dados_gerais={})            # licence identity missing — still fatal
    v, _ = check(no_id)
    assert v == FAILED, v

    odd_name = dict(good, file_name='licenca_renomeada_a_mao.pdf')  # non-SILIAMB filename
    v, reasons = check(odd_name)
    assert v == OK, (v, reasons)
    assert not any('expected LURH pattern' in r for r in reasons)  # no note about it anymore
    print('qa self-check OK')


if __name__ == '__main__':
    _demo()
