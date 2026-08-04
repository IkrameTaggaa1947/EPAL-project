"""Quality checks for a TUA extraction.

Returns a verdict — OK / REVIEW / FAILED — plus human-readable reasons, so a
non-technical operator can tell at a glance whether a certificate can be trusted
or needs a look.

The strongest check is free: the SILIAMB filename encodes the ground truth —
    Foios_TUA20230905002577_04.02.2024_a_03.02.2029.pdf
            └── Nº TUA ──┘  └ in-force ┘   └ validade ┘
so we can verify the extracted Nº TUA and both dates against the filename with
no external data. (This is exactly what would have caught the Aldeia da Ribeira
certificate whose whole dates section failed to extract.)
"""
import re
from datetime import datetime

# Nº TUA = 'TUA' + 14 digits; dates as DD.MM.YYYY separated by '_a_'
FNAME_RE = re.compile(r'(TUA\d{14})_(\d{2}\.\d{2}\.\d{4})_a_(\d{2}\.\d{2}\.\d{4})')
DATE_RE = re.compile(r'^\d{2}-\d{2}-\d{4}$')
VLE_RE = re.compile(r'^\s*\d+([.,]\d+)?(\s*[-a]\s*\d+([.,]\d+)?)?\s*$')

OK, REVIEW, FAILED = 'OK', 'REVIEW', 'FAILED'


def from_filename(name):
    """Ground-truth (Nº TUA, in-force date, validade) parsed from the filename,
    or None if the filename isn't in the expected SILIAMB pattern."""
    m = FNAME_RE.search(name or '')
    if not m:
        return None
    tua, vigor, val = m.groups()
    return {'Nº TUA': tua,
            'Data de Entrada em Vigor': vigor.replace('.', '-'),
            'Data de Validade': val.replace('.', '-')}


def _num(s):
    try:
        return float(str(s).replace(',', '.'))
    except (TypeError, ValueError):
        return None


def check(d):
    """Return (verdict, reasons). reasons is a list of '[LEVEL] message'."""
    g, e, loc = d['dados_gerais'], d['enquadramento'], d['localizacao']
    reasons = []

    def flag(level, msg):
        reasons.append((level, msg))

    # --- hard failures ------------------------------------------------------
    if not d.get('condicoes_rejeicao'):
        flag(FAILED, 'no discharge conditions (8.3.13) were extracted')
    if not g.get('Nº TUA'):
        flag(FAILED, 'Nº TUA missing')

    # --- filename cross-check (ground truth) --------------------------------
    fn = from_filename(d.get('file_name', ''))
    if fn:
        if g.get('Nº TUA') and g['Nº TUA'] != fn['Nº TUA']:
            flag(FAILED, f"Nº TUA mismatch: extracted {g['Nº TUA']} vs filename {fn['Nº TUA']}")
        for k in ('Data de Entrada em Vigor', 'Data de Validade'):
            got, exp = e.get(k), fn[k]
            if not got:
                flag(REVIEW, f"{k} missing (filename says {exp})")
            elif got != exp:
                flag(REVIEW, f"{k} {got} does not match filename {exp}")
    else:
        flag(REVIEW, 'filename not in the expected TUA pattern — cannot cross-check dates')

    # --- date validity ------------------------------------------------------
    for k in ('Data de Emissão', 'Data de Entrada em Vigor', 'Data de Validade'):
        v = e.get(k)
        if v and not DATE_RE.match(v):
            flag(REVIEW, f"{k} is not a valid date: {v!r}")

    # --- sections present ---------------------------------------------------
    if not d.get('autocontrolo'):
        flag(REVIEW, 'no self-monitoring rows (8.3.16) extracted')

    # --- coordinates inside mainland Portugal bounds ------------------------
    lon, lat = _num(loc.get('Longitude')), _num(loc.get('Latitude'))
    if lon is not None and not (-10.0 <= lon <= -6.0):
        flag(REVIEW, f'longitude out of expected range: {lon}')
    if lat is not None and not (36.5 <= lat <= 42.5):
        flag(REVIEW, f'latitude out of expected range: {lat}')

    # --- VLE sanity ---------------------------------------------------------
    for c in d.get('condicoes_rejeicao', []):
        vle = (c.get('VLE') or '').replace(' a ', '-')
        if vle and not VLE_RE.match(vle):
            flag(REVIEW, f"unusual VLE for {c.get('Parâmetro', '?')[:24]}: {c.get('VLE')!r}")

    # --- carry the extractor's own warnings (e.g. garbled param) ------------
    for w in d.get('_warnings', []):
        if 'garbled' in w or 'fragment' in w:
            flag(REVIEW, w)

    verdict = OK
    if any(l == REVIEW for l, _ in reasons):
        verdict = REVIEW
    if any(l == FAILED for l, _ in reasons):
        verdict = FAILED
    return verdict, [f'[{l}] {m}' for l, m in reasons]


# ------------------------------------------------------------------ self-check
def _demo():
    good = {'file_name': 'Foios_TUA20230905002577_04.02.2024_a_03.02.2029.pdf',
            'dados_gerais': {'Nº TUA': 'TUA20230905002577'},
            'enquadramento': {'Data de Emissão': '05-09-2023',
                              'Data de Entrada em Vigor': '04-02-2024',
                              'Data de Validade': '03-02-2029'},
            'localizacao': {'Longitude': '-6,89', 'Latitude': '40,28'},
            'condicoes_rejeicao': [{'Parâmetro': 'pH', 'VLE': '6-9'}],
            'autocontrolo': [{'x': 1}], '_warnings': []}
    v, _ = check(good)
    assert v == OK, v

    bad = dict(good, enquadramento={})            # dates section failed (Aldeia case)
    v, reasons = check(bad)
    assert v == REVIEW, (v, reasons)
    assert any('Validade' in r for r in reasons)

    broken = dict(good, condicoes_rejeicao=[])    # no conditions
    v, _ = check(broken)
    assert v == FAILED, v
    print('qa self-check OK')


if __name__ == '__main__':
    _demo()
