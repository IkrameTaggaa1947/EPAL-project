# -*- coding: utf-8 -*-
"""Extract the *diplomas* (legal instruments) cited in a `Legislação aplicável`
or `Avaliação da conformidade` text.

Both licence styles (TUA T-codes and LURH (a)/(1) markers) are already resolved
to inline `... (texto)` fields by the extractors, so this parser runs on that
free text and returns the diplomas cited, each PRIMARY diploma carrying any
amendments ("com as alterações introduzidas pelos Decretos-Leis ...") separately.

    extract("... artigo 53.º da Lei nº 58/2005 ... conjugada com o Decreto-Lei "
            "nº 152/97 ... com as alterações introduzidas pelos Decretos-Leis "
            "n.os 348/98 ... e 198/2008 ...")
    -> [{'diploma': 'Lei 58/2005', 'num': '58/2005', 'tipo': 'Lei', 'alteracoes': []},
        {'diploma': 'DL 152/97', 'num': '152/97', 'tipo': 'Decreto-Lei',
         'alteracoes': ['DL 348/98', 'DL 149/2004', 'DL 198/2008']}]

Noise (e.g. "a", "1", "Época Balnear", "Valor definido tendo em vista o bom
estado da massa de água ...") returns [].
"""
import re

MONTHS = ("janeiro|fevereiro|marco|março|abril|maio|junho|julho|agosto|"
          "setembro|outubro|novembro|dezembro")
TIPO = r"(Decreto[-\s]?Lei|Decreto\s+Regulamentar|Portaria|Resolu\w+|Lei)"
NUM = r"(\d+\s*/\s*\d{2,4})"
DATE = rf"(?:\s*,?\s*de\s+\d{{1,2}}\s*(?:\.|\s)\s*de\s+(?:{MONTHS}))?"

DIPLOMA_RE = re.compile(rf"{TIPO}\s*n?[.º°o\s]*{NUM}{DATE}", re.IGNORECASE)
# An amendment clause runs from "com as alterações introduzidas pel..." until the
# next of: a "(Quadro ...)"; a NEW primary introduced by "e o artigo"/"e o
# Decreto-Lei n.º<digit>"; a new sentence (". " + capital/paren); or end of text.
AMEND_RE = re.compile(
    r"com\s+as\s+altera\w+\s+introduzidas\s+pel\w+\s+(.*?)"
    r"(?:\(Quadro[^)]*\)?|,?\s+e\s+o\s+artigo\b"
    r"|,?\s+e\s+o\s+Decreto[-\s]?Lei\s+n[.º°o\s]*\d|\.\s+[A-Z(]|$)",
    re.IGNORECASE | re.DOTALL)
BARE_NUM = re.compile(r"\d+\s*/\s*\d{2,4}")


def _short(tipo, num):
    t = tipo.lower().replace(" ", "").replace("-", "")
    num = re.sub(r"\s+", "", num)
    if t.startswith("decretolei"):
        return f"DL {num}"
    if t.startswith("decretoregulamentar"):
        return f"DR {num}"
    if t.startswith("portaria"):
        return f"Portaria {num}"
    if t.startswith("resolu"):
        return f"Resolução {num}"
    if t == "lei":
        return f"Lei {num}"
    return f"{tipo} {num}"


def _canon_tipo(tipo):
    t = tipo.lower().replace(" ", "").replace("-", "")
    if t.startswith("decretolei"):
        return "Decreto-Lei"
    if t.startswith("decretoregulamentar"):
        return "Decreto Regulamentar"
    if t.startswith("portaria"):
        return "Portaria"
    if t.startswith("resolu"):
        return "Resolução"
    return "Lei"


def extract(text):
    """Return a list of primary-diploma dicts (see module docstring)."""
    if not text or not text.strip():
        return []
    amend_spans = []                       # (start, end, [bare nums])
    for m in AMEND_RE.finditer(text):
        nums = [re.sub(r"\s+", "", x) for x in BARE_NUM.findall(m.group(1))]
        amend_spans.append((m.start(), m.end(), nums))

    def in_amend(pos):
        return any(s <= pos < e for s, e, _ in amend_spans)

    prims, seen = [], set()
    for m in DIPLOMA_RE.finditer(text):
        if in_amend(m.start()):            # this number belongs to an amendment list
            continue
        short = _short(m.group(1), m.group(2))
        if short in seen:                  # dedupe OCR-duplicated citations
            continue
        seen.add(short)
        prims.append({"diploma": short, "num": re.sub(r"\s+", "", m.group(2)),
                      "tipo": _canon_tipo(m.group(1)), "alteracoes": [],
                      "_pos": m.start()})
    # attach each amendment clause's numbers to the nearest primary before it
    for s, e, nums in amend_spans:
        parent = None
        for p in prims:
            if p["_pos"] < s and (parent is None or p["_pos"] > parent["_pos"]):
                parent = p
        tgt = parent if parent is not None else (prims[0] if prims else None)
        if tgt is None:
            continue
        for n in nums:
            short = f"DL {n}"
            if short != tgt["diploma"] and short not in tgt["alteracoes"]:
                tgt["alteracoes"].append(short)
    for p in prims:
        p.pop("_pos", None)
    return prims


def diplomas_flat(text):
    """Convenience: 'DL 236/98 | DL 152/97' (primaries only, for a display cell)."""
    return " | ".join(d["diploma"] for d in extract(text))


if __name__ == "__main__":
    import sys
    print(extract(" ".join(sys.argv[1:])))
