# -*- coding: utf-8 -*-
"""Chave_ETAR — the join key between the licences and the ETAR asset registry.

The licences carry no asset code: a TUA/LURH names its installation in free text
('ETAR de Aldeia Velha'), while TabelaCentroide_ETAR names the same asset as
'ETAR Aldeia Velha' and holds CODMAXIMO, Centro_Operacional and the rest of the
registry attributes. The only thing the two sides share is that name, so the key
is the name folded to a comparable form:

    accents stripped -> upper case -> punctuation dropped
    -> Portuguese articles (de/da/do/das/dos) dropped -> spaces removed

Article folding is what takes the match from 157/269 to 227/269 licences; the
remaining ~42 are genuine name divergences ('ETAR Chãos', 'Carapito',
'ETAR de Campo Maior - Bacia A') and need the manual overrides below.

Computed HERE, in Python, and written into both feeds — not in Power Query,
which has no accent folding.
"""
import re
import unicodedata

ARTIGOS = {'DE', 'DA', 'DO', 'DAS', 'DOS', 'D'}

# licence name (as Chave_ETAR would fold it) -> registry name (same folding).
# Add a line here whenever QA reports a licence with no Centro Operacional and
# you have confirmed which asset it is. Keys/values are raw text; both sides go
# through chave_etar() below, so write them the way they appear in the file.
OVERRIDES = {
    # 'ETAR Chãos': 'ETAR de Chãos',
}


def chave_etar(nome):
    """Fold an ETAR name into the join key. Empty/None -> '' (stays unmatched)."""
    if nome is None:
        return ''
    txt = str(nome)
    for origem, destino in OVERRIDES.items():
        if _fold(origem) == _fold(txt):
            txt = destino
            break
    return _fold(txt)


def _fold(txt):
    sem_acentos = unicodedata.normalize('NFKD', str(txt)).encode('ascii', 'ignore').decode()
    limpo = re.sub(r'[^A-Z0-9 ]', ' ', sem_acentos.upper())
    return ''.join(t for t in limpo.split() if t not in ARTIGOS)
