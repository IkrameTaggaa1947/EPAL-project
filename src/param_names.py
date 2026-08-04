# -*- coding: utf-8 -*-
"""Single source of truth for parameter naming in the Power BI feed.

A parameter as written in a licence mixes three pieces of information:

    "Carência Bioquímica de Oxigénio (período de estiagem) (mg/L O2)"
     └── what is measured ──────────┘ └── when ───────────┘ └ unit ┘

Every generator splits it the same way, so the Excel files handed to Power BI
are already clean and Power Query only has to read and type them:

    Parametro  abbreviation only            'CBO5'
    Unidade    unit, without brackets       'mg/L O2'
    Periodo    'Normal' or 'Estiagem'       'Estiagem'

Import from the other build scripts:
    from param_names import split_parameter, abbreviate
"""
import re
import unicodedata

# Abbreviation table. Keys are accent-insensitive, lowercase, punctuation-free
# fragments of the full name; the first match wins, so longer, more specific
# fragments must come first.
_ABBREV = [
    ('carenciabioquimicadeoxigenio', 'CBO5'),
    ('carenciaquimicadeoxigenio', 'CQO'),
    ('totaldeparticulassolidasemsuspensao', 'SST'),
    ('solidossuspensostotais', 'SST'),
    ('oleosegorduras', 'O&G'),
    ('azotoamoniacal', 'N-NH4'),
    ('azotokjeldahl', 'N-Kjeldahl'),
    ('azototal', 'Nt'),
    ('azototatotal', 'Nt'),
    ('azototot', 'Nt'),
    ('fosforototal', 'Pt'),
    ('escherichiacoli', 'E. coli'),
    ('enterococosintestinais', 'Enterococos'),
    ('enterecocosintestinais', 'Enterococos'),
    ('coliformesfecais', 'CF'),
    ('nitratos', 'NO3'),
    ('nitritos', 'NO2'),
    ('fosfatos', 'PO4'),
    ('fenois', 'Fenóis'),
    ('cromiototal', 'Cr total'),
    ('detergentes', 'Detergentes'),
    ('oxigenio', 'O2'),
    ('caudal', 'Caudal'),
    ('ph', 'pH'),
]


def _key(text):
    """Accent-, case- and punctuation-insensitive key used for lookups."""
    text = unicodedata.normalize('NFKD', str(text or ''))
    text = text.encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+', '', text)


def abbreviate(name):
    """Full parameter name -> short code ('CBO5'). Unknown names are kept as-is."""
    k = _key(name)
    if not k:
        return ''
    for fragment, short in _ABBREV:
        if fragment in k:
            return short
    return re.sub(r'\s+', ' ', str(name)).strip()


def normalise_unit(unit):
    """Canonical unit spelling — licences mix 'mg/l' and 'mg/L', 'ml' and 'mL'."""
    u = re.sub(r'\s+', ' ', str(unit or '')).strip()
    if not u:
        return ''
    u = re.sub(r'\bmg/l\b', 'mg/L', u, flags=re.I)
    u = re.sub(r'\bml\b', 'mL', u, flags=re.I)
    u = re.sub(r'\bufc\b', 'ufc', u, flags=re.I)
    u = re.sub(r'\bm3\b', 'm³', u, flags=re.I)
    return u


def split_parameter(raw):
    """'<name> (período de estiagem) (<unit>)' -> (abbrev, unit, period).

    The unit is taken from the LAST bracket group, so a parameter that carries
    both an estiagem marker and a unit keeps the right one — the naive split on
    the first '(' returned 'período de estiagem (mg/L N' instead of 'mg/L N'.
    """
    text = re.sub(r'\s+', ' ', str(raw or '')).strip()
    if not text:
        return '', '', 'Normal'

    period = 'Estiagem' if 'estiagem' in _key(text) else 'Normal'
    # drop the estiagem marker wherever it sits, it is not part of the name
    text = re.sub(r'\(\s*per[íi]odo de estiagem\s*\)', ' ', text, flags=re.I)

    groups = re.findall(r'\(([^()]*)\)', text)
    unit = normalise_unit(groups[-1]) if groups else ''
    name = re.sub(r'\([^()]*\)', ' ', text)
    name = re.sub(r'\s+', ' ', name).strip(' -–—')

    return abbreviate(name), unit, period


def clean_vle(value):
    """Normalise a limit value: '6,0 a 9,0' -> '6,0 - 9,0'.

    Only the range keyword is rewritten. The previous Power Query step replaced
    EVERY letter 'a' with '-', which silently corrupted any textual limit.
    """
    text = re.sub(r'\s+', ' ', str(value or '')).strip()
    if not text:
        return ''
    return re.sub(r'(?<=[\d,.])\s*a\s*(?=[\d,.])', ' - ', text)
