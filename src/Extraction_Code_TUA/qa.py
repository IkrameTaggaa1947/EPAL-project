"""Quality checks for a TUA extraction.

Returns a verdict — OK / FAILED — plus human-readable reasons. The REVIEW verdict
was removed: only a genuine FAILED (no conditions extracted, Nº TUA missing/mismatch)
keeps a certificate out of the master. All the other checks still run and are
returned as informational notes (logged + in qa_report.xlsx); they no longer
divert files to review/.

The strongest check is free: the SILIAMB filename encodes the ground truth —
    Foios_TUA20230905002577_04.02.2024_a_03.02.2029.pdf
            └── Nº TUA ──┘  └ in-force ┘   └ validade ┘
so we can verify the extracted Nº TUA and both dates against the filename with
no external data. (This is exactly what would have caught the Aldeia da Ribeira
certificate whose whole dates section failed to extract.)
"""
import re
from datetime import datetime

# Nº TUA = 'TUA' + 14 digitos, onde quer que apareca no nome.
TUA_RE = re.compile(r'(TUA\d{14})(?!\d)')   # sem \b final: falha antes de '_'
# Uma data dentro do nome: DD.MM.AAAA ou DD-MM-AAAA.
FNAME_DATE = re.compile(r'(\d{2})[.\-](\d{2})[.\-](\d{4})')
DATE_RE = re.compile(r'^\d{2}-\d{2}-\d{4}$')
VLE_RE = re.compile(r'^\s*\d+([.,]\d+)?(\s*[-a]\s*\d+([.,]\d+)?)?\s*$')

OK, REVIEW, FAILED = 'OK', 'REVIEW', 'FAILED'


def datas_do_nome(name):
    """As datas plausiveis que aparecem no nome do ficheiro, por ordem.

    Nao ha UM formato: os nomes vindos do SILIAMB, os renomeados a mao e os das
    renovacoes escrevem-nas de maneiras diferentes --
        Foios_TUA20230905002577_04.02.2024_a_03.02.2029.pdf
        Cardigos TUA20230130000334 (27.01.2023 a 26.01.2026).pdf
        TUA20230130000334_ETAR_Cardigos_27.01.2026_26.01.2031.pdf
        Vales de Cardigos_PIP004612.2017.RH5A (03-04-2017 a 03-04-2018).pdf
    -- por isso procuram-se DATAS, nao um padrao de nome. Valida-se dia, mes e
    ano para nao apanhar pedacos de codigos como 'PIP004612.2017.RH5A'.
    """
    saida = []
    for d, m, a in FNAME_DATE.findall(name or ''):
        if 1 <= int(d) <= 31 and 1 <= int(m) <= 12 and 1900 <= int(a) <= 2100:
            saida.append('%s-%s-%s' % (d, m, a))
    return saida


def from_filename(name):
    """O que o NOME do ficheiro diz: Nº TUA e/ou o par de datas.

    Devolve so o que conseguiu ler (pode ser so o Nº TUA, so as datas, ou
    nada). Um nome fora do padrao NAO e um problema -- e so menos uma fonte
    para confirmar o que foi extraido do documento."""
    out = {}
    m = TUA_RE.search(name or '')
    if m:
        out['Nº TUA'] = m.group(1)
    datas = datas_do_nome(name)
    if len(datas) >= 2:
        out['Data de Entrada em Vigor'] = datas[0]
        out['Data de Validade'] = datas[-1]
    return out


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

    # --- o que o nome do ficheiro confirma ----------------------------------
    # Um nome fora do padrao NAO e motivo de aviso: ha varias convencoes em uso
    # e nenhuma delas e errada. So se compara o que o nome REALMENTE traz.
    fn = from_filename(d.get('file_name', ''))
    if fn.get('Nº TUA') and g.get('Nº TUA') and g['Nº TUA'] != fn['Nº TUA']:
        flag(FAILED, f"Nº TUA mismatch: extracted {g['Nº TUA']} vs filename {fn['Nº TUA']}")
    for k in ('Data de Entrada em Vigor', 'Data de Validade'):
        exp = fn.get(k)
        if not exp:
            continue                      # o nome nao diz nada sobre esta data
        got = e.get(k)
        if got and got != exp:
            flag(REVIEW, f"{k} {got} does not match filename {exp}")

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

    # REVIEW step removed: only a genuine FAILED keeps a certificate out of the
    # master. All the lower-level checks still run and are returned as informational
    # notes (logged + in qa_report.xlsx), they just no longer divert files to review/.
    verdict = FAILED if any(l == FAILED for l, _ in reasons) else OK
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
    assert v == OK, (v, reasons)                  # REVIEW step removed -> not blocked

    # Um nome fora do padrao NAO gera aviso nenhum: ha varias convencoes em uso.
    outro = dict(good, file_name='TUA20230130000334_ETAR_Cardigos_27.01.2026_26.01.2031.pdf',
                 dados_gerais={'Nº TUA': 'TUA20230130000334'},
                 enquadramento={'Data de Entrada em Vigor': '27-01-2026',
                                'Data de Validade': '26-01-2031'})
    v, reasons = check(outro)
    assert v == OK, (v, reasons)
    assert not any('pattern' in r for r in reasons), reasons

    # O nome le-se em qualquer das convencoes em uso.
    assert from_filename('Foios_TUA20230905002577_04.02.2024_a_03.02.2029.pdf') == {
        'Nº TUA': 'TUA20230905002577',
        'Data de Entrada em Vigor': '04-02-2024',
        'Data de Validade': '03-02-2029'}
    assert from_filename('Cardigos TUA20230130000334 (27.01.2023 a 26.01.2026).pdf')[
        'Data de Validade'] == '26-01-2026'
    # um codigo NAO e uma data
    assert from_filename('Fronteira_2012.000813.000.T.L.RJ.DAR.pdf') == {}

    broken = dict(good, condicoes_rejeicao=[])    # no conditions
    v, _ = check(broken)
    assert v == FAILED, v
    print('qa self-check OK')


if __name__ == '__main__':
    _demo()
