"""
Reliable LURH extractor (Licença de Utilização dos Recursos Hídricos —
Rejeição de Águas Residuais, SILIAMB/APA).

Strategy (the same philosophy as the TUA extractor — resilient to layout
shifts, and never based on x/y positions of words):
  * Values are located by their LABEL text ("Código APA", "Massa de água", ...),
    matched accent/case-insensitively at the start of a line, never by position.
  * Sections are located by their MARKER text ("Caracterização da rejeição",
    "Condições de descarga das águas residuais...", "Autocontrolo", ...), so the
    "no ano de arranque" VLE table is kept apart from the "condições normais" one.
  * Parameter tables are read twice, and the best result wins:
      1) PyMuPDF `find_tables()` with columns mapped by HEADER NAME (the LURH
         PDFs are ruled, exactly like the TUA ones), and
      2) a text-layer parser driven by the tables' small closed vocabularies
         (Entrada/Saída, Mensal/Quinzenal/..., Pontual/Composta), used whenever
         find_tables yields nothing for a section.
  * Wrapped lines are rejoined ("Carência Bioquímica de / Oxigénio" ->
    "Carência Bioquímica de Oxigénio").

Run:  python extract_lurh.py [path/to/LURH.pdf]
Outputs (in the configured output dir): <name>_extracted.json + <name>_extracted.xlsx
"""
import fitz
import re
import os
import sys
import json
import shutil
import unicodedata
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

LIC_RE = re.compile(r'\bL\d{6}\.\d{4}\.RH\w+\b')
NUM = r'\d+(?:[.,]\d+)?'
RANGE = rf'{NUM}(?:\s*(?:-|a)\s*{NUM})?'          # "25", "6-9", "6,0 a 9,0"
CODE = r'\(?[a-z]\)'                              # "(a)" or the occasional "a)"

FREQ = (r'Diária|Diário|Semanal|Quinzenal|Mensal|Bimensal|Bimestral|Trimestral|'
        r'Quadrimestral|Semestral|Anual|Contínua|Contínuo')
TIPO = r'Pontual|Composta(?:\s*\((?:i{1,3}|iv|v)\))?|Em\s+cont[íi]nuo|Contínuo|Contínua|Registo'
AUTO_END = re.compile(rf'(?P<freq>{FREQ})\s+(?P<tipo>{TIPO})\s*$')
AUTO_LOCAL = re.compile(r'^(Entrada|Sa[ií]da)\b')
MON_LOCAL = re.compile(r'^(P1 e P2|P[0-9]|Montante|Jusante)\b')
MON_FREQ = re.compile(rf'^(?:{FREQ})$')

BOILER = ('pag.', 'siliamb', 'codigo documento', 'codigo verifica')


# ----------------------------------------------------------------- helpers
def norm(s):
    if not s:
        return ''
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().lower()


def _key(s):
    return norm(s).replace(' ', '')


CTRL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')   # PDF artifacts (e.g. 'Decreto\x02Lei')


def clean_text(s):
    s = CTRL_RE.sub('-', s or '')          # the corpus uses \x02 as a hyphen
    return re.sub(r'\s+', ' ', s).strip()


def fold_words(s):
    """Accent/case-folded words with '*' markers dropped — the unit of
    label matching (labels are matched word-by-word, never by position)."""
    return [w for w in norm(s.replace('*', ' ')).split() if w]


def date_norm(s):
    """'2022/10/11' or '11-10-2022' -> '11-10-2022'."""
    s = clean_text(s)
    m = re.search(r'(\d{4})/(\d{2})/(\d{2})', s)
    if m:
        return f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
    m = re.search(r'(\d{2})-(\d{2})-(\d{4})', s)
    return m.group(0) if m else s


def checkbox(s):
    t = clean_text(s)
    if re.search(r'\|\s*[Xx☒]\s*\||^[Xx☒]$', t):
        return 'Sim'
    if re.search(r'\|\s*_\s*\||^_$', t):
        return 'Não'
    return t


def dewrap(txt):
    """Join the physical lines of a table cell (trailing-space = word boundary),
    exactly as in the TUA extractor."""
    if not txt:
        return ''
    out, prev_boundary = '', True
    for ln in txt.split('\n'):
        boundary = ln != ln.rstrip()
        seg = ln.strip()
        if not seg:
            continue
        out = seg if not out else (out + ' ' + seg if prev_boundary else out + seg)
        prev_boundary = boundary
    return re.sub(r'\s+', ' ', out).strip()


# ----------------------------------------------------------------- document IO
def load_pdf(pdf_path):
    """(page_texts, table_rows). page_texts is the text layer in natural reading
    order; table_rows is every ruled-table row on every page as de-wrapped cell
    strings (may be empty if the PDF has no detectable ruling)."""
    doc = fitz.open(pdf_path)
    pages, rows = [], []
    for page in doc:
        pages.append(page.get_text('text'))
        try:
            for t in page.find_tables().tables:
                for trow in t.rows:
                    rows.append([dewrap(page.get_textbox(b)) if b else '' for b in trow.cells])
        except Exception:
            pass                       # ruling detection is an overlay, never fatal
    doc.close()
    return pages, rows


def doc_lines(pages):
    """All lines of the document with page footers stripped:
    'N/M -' page marks and the bare licence-code line under them."""
    out = []
    for p in pages:
        for ln in CTRL_RE.sub('-', p.replace('\r', '')).split('\n'):
            s = ln.strip()
            if not s:
                continue
            if re.fullmatch(r'\d+\s*/\s*\d+\s*-?', s) or re.fullmatch(LIC_RE, s):
                continue
            if any(b in norm(s) for b in BOILER):
                continue
            out.append(s)
    return out


# ----------------------------------------------------------------- label engine
def match_label(line, label_words_list):
    """If `line` starts with one of the known labels (accent/case/'*'-insensitive,
    matched word by word), return (label_index, value_after_label) else None.
    Longest label wins, so 'Denominação do meio recetor' beats 'Meio Recetor'."""
    words = line.split()
    fw = fold_words(line)
    best = None
    for idx, lw in label_words_list:
        k = len(lw)
        if len(fw) >= k and fw[:k] == lw:
            if best is None or k > best[1]:
                best = (idx, k)
    if best is None:
        return None
    idx, k = best
    # the label consumed k folded words == k raw words ('*' stays glued to its word)
    return idx, ' '.join(words[k:]).strip()


def _prep_labels(labels):
    """[(canonical_name, [written forms...])] -> matcher list [(i, fold_words)]."""
    out = []
    for i, (_, forms) in enumerate(labels):
        for f in forms:
            out.append((i, fold_words(f)))
    return out


def parse_labels(lines, labels, stop_fold=(), strip_prefixes=None):
    """Generic label/value reader for a section given as text lines.
    A line that starts with a known label opens that field; following lines that
    match no label are appended to the open field (this is how wrapped values —
    and wrapped label tails like '(e.p)' — are absorbed; strip_prefixes then
    removes known label-tail phrases from the value)."""
    matcher = _prep_labels(labels)
    out, cur = {}, None
    for ln in lines:
        nf = norm(ln)
        if any(nf.startswith(s) for s in stop_fold):
            break
        m = match_label(ln, matcher)
        if m is not None:
            idx, val = m
            name = labels[idx][0]
            if name not in out or out[name] == '':
                out[name] = val
                cur = name
            else:
                cur = None                    # repeated label: keep the first hit
        elif cur is not None:
            out[cur] = (out[cur] + ' ' + ln).strip()
    for name, phrases in (strip_prefixes or {}).items():
        v = out.get(name, '')
        for ph in phrases:
            pw = fold_words(ph)
            vw = v.split()
            if fold_words(v)[:len(pw)] == pw:
                v = ' '.join(vw[len(pw):])
        out[name] = clean_text(v)
    return {k: clean_text(v) for k, v in out.items()}


def section(lines, start_marker, end_markers, include_start_line=False):
    """The slice of lines between a start marker and the first end marker.
    Markers are matched as accent-insensitive prefixes of a line."""
    s = norm(start_marker)
    ends = [norm(e) for e in end_markers]
    i = next((i for i, ln in enumerate(lines) if norm(ln).startswith(s)), None)
    if i is None:
        return []
    j = next((j for j in range(i + 1, len(lines))
              if any(norm(lines[j]).startswith(e) for e in ends)), len(lines))
    return lines[i if include_start_line else i + 1:j]


# ----------------------------------------------------------------- sections
HEADER_LABELS = [
    ('Nº Processo', ['Processo n.º:', 'Processo nº:', 'Processo n.o:']),
    ('Nº Licença', ['Utilização n.º:', 'Utilizacao n.º:', 'Utilização nº:', 'Autorização n.º:']),
    ('Data de Início', ['Início:', 'Inicio:']),
    ('Data de Validade', ['Validade:']),
]

IDENT_LABELS = [
    ('Código APA', ['Código APA']),
    ('País', ['País']),
    ('NIF', ['Número de Identificação Fiscal', 'Número de Identificação fiscal']),
    ('Requerente', ['Nome/Denominação Social']),
    ('Idioma', ['Idioma']),
    ('Morada', ['Morada']),
    ('Localidade', ['Localidade']),
    ('Código Postal', ['Código Postal']),
    ('Concelho', ['Concelho']),
    ('Telefones', ['Telefones']),
    ('Fax', ['Fax']),
    ('Obrigação de correcção de Dados de Perfil', ['Obrigação de correcção de Dados de Perfil',
                                                   'Obrigação de correção de Dados de Perfil']),
]

TRAT_LABELS = [
    ('Designação', ['Designação']),
    ('Nível de tratamento', ['Nível de tratamento implementado']),
    ('Tipo de tratamento', ['Tipo de tratamento']),
    ('Caudal Máximo descarga', ['Caudal Máximo descarga', 'Caudal máximo descarga']),
    ('Caudal Médio descarga', ['Caudal Médio descarga', 'Caudal médio descarga']),
    ('Nut III – Concelho – Freguesia', ['Nut III – Concelho – Freguesia', 'Nut III - Concelho - Freguesia']),
    ('Longitude', ['Longitude']),
    ('Latitude', ['Latitude']),
    ('Ano de arranque', ['Ano de arranque']),
    ('População servida (e.p.)', ['População servida (e.p.)', 'População servida (e.p)']),
    ('Ano horizonte de projeto', ['Ano horizonte de projeto']),
    ('População servida no horizonte',
     ['População servida no ano horizonte de projeto (e.p)',
      'População servida no ano horizonte de projeto']),
]

REJ_LABELS = [
    ('Origem das águas residuais', ['Origem das águas residuais']),
    ('Designação da rejeição', ['Designação da rejeição']),
    ('Meio Recetor', ['Meio Recetor', 'Meio Receptor']),
    ('Margem', ['Margem']),
    ('Denominação do meio recetor', ['Denominação do meio recetor', 'Denominação do meio receptor']),
    ('Sistema de Descarga', ['Sistema de Descarga']),
    ('Valorização ou reutilização', ['Valorização ou reutilização']),
    ('Caudal Reutilizado', ['Caudal Reutilizado']),
    ('Finalidades Efluente', ['Finalidades Efluente']),
    ('Nut III – Concelho – Freguesia', ['Nut III – Concelho – Freguesia', 'Nut III - Concelho - Freguesia']),
    ('Longitude', ['Longitude']),
    ('Latitude', ['Latitude']),
    ('Região Hidrográfica', ['Região Hidrográfica']),
    ('Bacia Hidrográfica', ['Bacia Hidrográfica']),
    ('Sub-Bacia Hidrográfica', ['Sub-Bacia Hidrográfica', 'Sub-bacia Hidrográfica']),
    ('Tipo de massa de água', ['Tipo de massa de água']),
    ('Massa de água', ['Massa de água', 'Massa de Água']),
    ('Classificação da massa de água', ['Classificação do estado/potencial ecológico']),
]

SECTION_STOPS = ['identificacao', 'caracterizacao do', 'caracterizacao da rejeicao',
                 'condicoes gerais', 'condicoes especificas', 'outras condicoes', 'anexos']


def dados_gerais(lines):
    head = section(lines, 'Processo n.', ['Identificação', 'Caracterização do'], include_start_line=True)
    out = parse_labels(head, HEADER_LABELS, stop_fold=('licenca de utilizacao',))
    ident = section(lines, 'Identificação', ['Caracterização do'])
    out.update(parse_labels(ident, IDENT_LABELS))
    for k in ('Data de Início', 'Data de Validade'):
        if out.get(k):
            out[k] = date_norm(out[k])
    if 'Obrigação de correcção de Dados de Perfil' in out:
        out['Obrigação de correcção de Dados de Perfil'] = checkbox(out['Obrigação de correcção de Dados de Perfil'])
    if out.get('Nº Licença'):
        m = LIC_RE.search(out['Nº Licença'].replace(' ', ''))
        if m:
            out['Nº Licença'] = m.group(0)
    return out


def tratamento(lines):
    seg = section(lines, 'Caracterização do', ['Caracterização da rejeição', 'Condições Gerais'])
    out = parse_labels(seg, TRAT_LABELS,
                       strip_prefixes={'População servida no horizonte': ['(e.p)', '(e.p.)']})
    return out


AFL_VOL = re.compile(rf'^(Volume\s+(?:M[áa]ximo|M[ée]dio)\s+mensal)\s+({NUM})?\s*\(m3\)\s*$', re.I)
AFL_PAR = re.compile(rf'^(CBO5|CQO|N|P)\s+(?:({NUM})\s+)?\(mg/L[^)]*\)\s*$')
AFL_VOL_L = re.compile(r'^Volume\s+(M[áa]ximo|M[ée]dio)\s+mensal\s*$', re.I)   # label alone…
AFL_PAR_L = re.compile(r'^(CBO5|CQO|N|P)\s*$')                                 # …value on next line
AFL_VAL = re.compile(rf'^({NUM})\s*\((?:m3|mg/L[^)]*)\)\s*$')


def rejeicao(lines):
    seg = section(lines, 'Caracterização da rejeição',
                  ['Condições Gerais', 'Condições Específicas', 'Anexos'])
    out = parse_labels(seg, REJ_LABELS,
                       strip_prefixes={'Classificação da massa de água':
                                       ['(superficial) ou estado (subterrânea) da massa de água',
                                        '(superficial) ou estado (subterranea) da massa de agua']})
    if 'Valorização ou reutilização' in out:
        out['Valorização ou reutilização'] = checkbox(out['Valorização ou reutilização'])
    # Características do Afluente Bruto — its labels (N, P, ...) are too short for
    # the generic engine, so they are read with their unit as the anchor. Both
    # layouts occur: 'CBO5 474.0 (mg/L O2)' on one line, or label / value lines.
    for k, ln in enumerate(seg):
        nxt = seg[k + 1] if k + 1 < len(seg) else ''
        m = AFL_VOL.match(ln)
        if m:
            out['Volume mensal afluente (m3)'] = m.group(2) or ''
            out['Volume mensal afluente (tipo)'] = 'Máximo' if 'x' in norm(m.group(1)) else 'Médio'
            continue
        m = AFL_VOL_L.match(ln)
        if m and AFL_VAL.match(nxt):
            out['Volume mensal afluente (m3)'] = AFL_VAL.match(nxt).group(1)
            out['Volume mensal afluente (tipo)'] = 'Máximo' if 'x' in norm(m.group(1)) else 'Médio'
            continue
        m = AFL_PAR.match(ln)
        if m:
            out[f'Afluente {m.group(1)} (mg/L)'] = m.group(2) or ''
            continue
        m = AFL_PAR_L.match(ln)
        if m and AFL_VAL.match(nxt):
            out[f'Afluente {m.group(1)} (mg/L)'] = AFL_VAL.match(nxt).group(1)
    return out


def condicoes_texto(lines):
    def items(seg):
        out, cur = [], None
        for ln in seg:
            m = re.match(r'^(\d+)ª\s+(.*)$', ln)
            if m:
                if cur:
                    out.append(clean_text(cur))
                cur = m.group(2)
            elif cur is not None:
                cur += ' ' + ln
        if cur:
            out.append(clean_text(cur))
        return out
    return {
        'Gerais': items(section(lines, 'Condições Gerais', ['Condições Específicas', 'Outras Condições', 'Anexos'])),
        'Específicas': items(section(lines, 'Condições Específicas', ['Outras Condições', 'Anexos'])),
        'Outras': items(section(lines, 'Outras Condições', ['Anexos'])),
    }


# ------------------------------------------------- VLE (condições de descarga)
VLE_ROW = re.compile(
    rf'^(?P<param>.+?)\s+(?P<v1>{RANGE})(?:\s+(?P<v2>{RANGE}))?\s*(?P<codes>(?:{CODE}\s*)+)$')
VLE_HDR = re.compile(r'^par[âa]metro\b.*\bvle\b.*legisla', re.I)
DL15297_PCT = {'70-90', '75', '90'}       # the DL 152/97 minimum-removal triple


def _regime(title):
    n = norm(title)
    if 'arranque' in n:
        return 'Ano de arranque'
    if 'normais' in n or not n:
        return 'Normal'
    return clean_text(title) or 'Normal'


def vle_from_lines(lines, warnings):
    """Every 'Condições de descarga das águas residuais …' block: its regime,
    its Observações paragraph, and its parameter rows. Handles both text
    layouts: cells joined per visual row, and one cell per line (real PDFs)."""
    conds, obs = [], {}
    idxs = [i for i, ln in enumerate(lines)
            if norm(ln).startswith('condicoes de descarga das aguas residuais')]
    STOP = ('legislacao', 'avaliacao de conformidade', 'programa de',
            'autocontrolo', 'condicoes de descarga')

    for bi, i in enumerate(idxs):
        end = idxs[bi + 1] if bi + 1 < len(idxs) else len(lines)
        regime = _regime(lines[i][len('Condições de descarga das águas residuais'):])
        seg = lines[i + 1:end]

        def emit(param, v1, v2, codes):
            param = clean_text(param)
            pct, vle = '', v1
            if v2:                                    # two values: % remoção + VLE
                pct, vle = v1, v2
            elif has_pct and v1 in DL15297_PCT and 'ph' not in norm(param):
                # single value under a 4-column header: the DL 152/97 removal triple
                pct, vle = v1, ''
                warnings.append(f"'{param}' ({regime}): single value {v1!r} read as "
                                f"% remoção (4-column table) — confirm against the PDF")
            conds.append({'Regime': regime, 'Parâmetro': param,
                          'VLE (% mín. remoção)': pct, 'VLE': vle,
                          'Legislação aplicável': codes})

        # ---- header: one joined line, or a run of bare header-cell lines ----
        hdr = next((k for k, ln in enumerate(seg) if VLE_HDR.match(norm(ln))), None)
        cellwise, body_start, has_pct = False, None, False
        if hdr is not None:
            has_pct = 'remo' in norm(seg[hdr]) or 'redu' in norm(seg[hdr])
            body_start = hdr + 1
        else:
            p = next((k for k, ln in enumerate(seg) if norm(ln) == 'parametro'), None)
            if p is not None:
                le = next((k for k in range(p + 1, min(p + 8, len(seg)))
                           if norm(seg[k]).startswith('legislacao aplicavel')), None)
                if le is not None:
                    hdr, cellwise, body_start = p, True, le + 1
                    has_pct = any('remo' in norm(seg[k]) or 'redu' in norm(seg[k])
                                  for k in range(p, le + 1))
        # Observações between the intro sentence and the table header
        oidx = next((k for k, ln in enumerate(seg[:hdr if hdr is not None else len(seg)])
                     if norm(ln) == 'observacoes'), None)
        if oidx is not None and hdr is not None:
            obs[regime] = clean_text(' '.join(seg[oidx + 1:hdr]))
        if hdr is None:
            continue

        if cellwise:
            # param line(s) -> value line(s) -> codes line, each cell on its own line
            param_buf, v1, v2 = [], None, None
            for ln in seg[body_start:]:
                n, s = norm(ln), ln.strip()
                if n.startswith(STOP):
                    if param_buf and v1 is not None:
                        emit(' '.join(param_buf), v1, v2, '')
                    break
                if param_buf and re.fullmatch(RANGE, s):
                    if v1 is None:
                        v1 = s
                    else:
                        v2 = s
                    continue
                if v1 is not None and re.fullmatch(rf'(?:{CODE}\s*)+', s):
                    emit(' '.join(param_buf), v1, v2, ' '.join(re.findall(CODE, s)))
                    param_buf, v1, v2 = [], None, None
                    continue
                if v1 is not None:            # new parameter began without a codes cell
                    emit(' '.join(param_buf), v1, v2, '')
                    param_buf, v1, v2 = [], None, None
                param_buf.append(ln)
            continue

        pending = ''
        for ln in seg[body_start:]:
            n = norm(ln)
            if n.startswith(STOP):
                break
            m = VLE_ROW.match(ln)
            if not m:
                pending = (pending + ' ' + ln).strip()   # wrapped parameter fragment
                continue
            emit((pending + ' ' + m.group('param')).strip(), m.group('v1'), m.group('v2'),
                 ' '.join(re.findall(CODE, m.group('codes'))))
            pending = ''
    return conds, obs


def vle_from_tables(rows):
    """The same blocks read from ruled tables: regime tracked from the preceding
    marker row, columns mapped on every header row by NAME."""
    conds, regime, ci = [], None, None
    for r in rows:
        joined = ' '.join(c for c in r if c)
        n = norm(joined)
        if n.startswith('condicoes de descarga das aguas residuais'):
            regime = _regime(joined[len('Condições de descarga das águas residuais'):])
            ci = None
            continue
        if n.startswith(('legislacao', 'avaliacao de conformidade', 'programa de', 'autocontrolo')):
            ci = None
            continue
        if regime and col_idx(r, 'parametro') is not None and col_exact_any(r, 'vle') is not None:
            ci = {'Parâmetro': col_idx(r, 'parametro'),
                  'VLE (% mín. remoção)': col_idx(r, 'remo', 'redu'),
                  'VLE': col_exact_any(r, 'vle'),
                  'Legislação aplicável': col_idx(r, 'legislacao aplicavel')}
            continue
        if ci is None:
            continue
        cell = lambda k: clean_text(r[ci[k]]) if ci.get(k) is not None and ci[k] < len(r) else ''
        param = cell('Parâmetro')
        if not param or not (cell('VLE') or cell('VLE (% mín. remoção)')):
            continue
        conds.append({'Regime': regime, 'Parâmetro': param,
                      'VLE (% mín. remoção)': cell('VLE (% mín. remoção)'),
                      'VLE': cell('VLE'),
                      'Legislação aplicável': ' '.join(re.findall(CODE, cell('Legislação aplicável')))
                                              or cell('Legislação aplicável')})
    return conds


def col_idx(header, *names):
    for i, c in enumerate(header):
        kc = _key(c)
        if kc and any(_key(n) in kc for n in names):
            return i
    return None


def col_exact_any(header, target):
    t = _key(target)
    for i, c in enumerate(header):
        if _key(c) == t:
            return i
    return None


# ------------------------------------------------------------------- legends
def legends(lines):
    """(legislacao_map, avaliacao_map, avaliacao_texto). Every 'Legislação' block
    contributes '(a)/(b)' letters; the Avaliação blocks contribute letters when
    present, plus their full text (some licences use no letters at all — the
    whole paragraph then applies to every parameter)."""
    legis, aval, aval_txt = {}, {}, []

    def collect(start_pred, end_prefixes):
        blocks, cur = [], None
        for ln in lines:
            n = norm(ln)
            if start_pred(n):
                cur = []
                continue
            if cur is not None:
                if any(n.startswith(e) for e in end_prefixes):
                    blocks.append(' '.join(cur))
                    cur = None
                else:
                    cur.append(ln)
        if cur:
            blocks.append(' '.join(cur))
        return blocks

    def letters(text, into):
        parts, cur = re.split(r'(\([a-z]\))', ' ' + text), None
        for p in parts:
            m = re.fullmatch(r'\(([a-z])\)', p.strip())
            if m:
                cur = m.group(1)
            elif cur is not None and p.strip():
                into[cur] = clean_text((into.get(cur, '') + ' ' + p))

    for b in collect(lambda n: n == 'legislacao',
                     ('avaliacao de conformidade', 'condicoes de descarga', 'programa de',
                      'autocontrolo', 'observacoes')):
        letters(b, legis)
    for b in collect(lambda n: n.startswith('avaliacao de conformidade'),
                     ('programa de', 'autocontrolo', 'condicoes de descarga', 'anexos')):
        aval_txt.append(clean_text(b))
        letters(b, aval)
    return legis, aval, clean_text(' '.join(dict.fromkeys(aval_txt)))


def resolve(codes, legis, aval, aval_texto):
    """A row's letter codes -> (legislação texto, avaliação texto). Letters live
    in whichever legend defined them; when the Avaliação legend has no letters,
    its whole text applies to every row."""
    ks = re.findall(r'\(?([a-z])\)', codes or '')
    lt = ' | '.join(legis[k] for k in ks if k in legis)
    at = ' | '.join(aval[k] for k in ks if k in aval)
    if not at and aval_texto:
        at = aval_texto
    return lt, at


# ----------------------------------------------- meio recetor + autocontrolo
MET_ANCHOR = re.compile(r'\b(?:Metodologia\b|Anexo\s+[IVX0-9]+\s+do\s+Decreto)')


def _method_split(buf):
    """A row buffer (list of lines) -> (before_method, method). The analytic-method
    column starts at 'Metodologia' or 'Anexo <n> do Decreto-Lei' (or is just '-')."""
    joined = ' '.join(buf)
    m = MET_ANCHOR.search(joined)
    if m:
        return clean_text(joined[:m.start()]), clean_text(joined[m.start():])
    m = re.search(r'\s-\s*$', joined) or re.search(r'\s-\s', joined)
    if m:
        return clean_text(joined[:m.start()]), '-'
    return clean_text(joined), ''


def monitorizacao_from_lines(lines):
    seg = section(lines, 'Programa de monitorização do meio recetor',
                  ['Autocontrolo', 'Anexos'])
    if not seg:
        return {}
    hdr = next((i for i, ln in enumerate(seg)
                if norm(ln).startswith('local ') and 'parametro' in norm(ln)), None)
    pontos, rows = {}, []
    for ln in seg[:hdr if hdr is not None else len(seg)]:
        m = re.match(r'^(P\d)\s*[:\-–]\s*(.+)$', ln)
        if m:
            pontos[m.group(1)] = clean_text(m.group(2))
    obs = clean_text(' '.join(ln for ln in seg[:hdr if hdr is not None else len(seg)]
                              if not re.match(r'^(P\d)\s*[:\-–]', ln)
                              and norm(ln) != 'observacoes'))
    if hdr is None:
        return {'Observações': obs, 'Pontos': pontos, 'rows': []}
    body, buf, local, freq, robs = seg[hdr + 1:], [], None, '', []
    pextra, met_open = [], False
    hdr_words = {'local', 'de', 'amostragem', 'parametro', 'metodo', 'analitico',
                 'frequencia', 'observacoes'}

    def flush():
        if local is None:
            return
        before, met = _method_split(buf)
        if pextra:
            before = clean_text(before + ' ' + ' '.join(pextra))
        # page-break repair: when freq arrived before the método finished, the
        # remainders of the parameter/método cells were collected after it.
        leftover = list(robs)
        if met and met != '-' and not met.rstrip().endswith('junho.'):
            if leftover:                    # leading parameter-cell remainder ('O2) ...')
                toks, moved = leftover[0].split(), []
                while toks and re.fullmatch(r'[^()]*\)', toks[0]):
                    moved.append(toks.pop(0))
                if moved:
                    before = clean_text(before + ' ' + ' '.join(moved))
                    if toks:
                        leftover[0] = ' '.join(toks)
                    else:
                        leftover.pop(0)
            while leftover and not re.match(r'^(aproximadamente|durante|metros|ponto|\()',
                                            leftover[0], re.I):
                met = clean_text(met + ' ' + leftover.pop(0))
                if met.rstrip().endswith('junho.'):
                    break
        rows.append({'Local': local, 'Parâmetro': clean_text(before),
                     'Método analítico': met, 'Frequência': freq,
                     'Observações': clean_text(' '.join(leftover))})

    for ln in body:
        n = norm(ln)
        if set(n.split()) <= hdr_words:
            continue                                   # wrapped header fragments
        lm = MON_LOCAL.match(ln)
        if lm and not re.match(r'^(P\d)\s*[:\-–]', ln):
            flush()
            local, freq, robs, pextra, met_open = lm.group(1), '', [], [], False
            buf = [ln[lm.end():].strip()]
            # page-break scramble: all columns' first lines glued on one line
            # ('P2 Enterococos Anexo I do Decreto-Lei n.º Mensal aproximadamente...')
            am = MET_ANCHOR.search(buf[0])
            fm = am and re.search(rf'\s({FREQ})\b', buf[0][am.start():])
            if fm:
                a = am.start()
                freq = fm.group(1)
                tail = buf[0][a + fm.end():].strip()
                robs = [tail] if tail else []
                buf = [buf[0][:a + fm.start()]]
                met_open = not buf[0].rstrip().endswith('junho.')
            continue
        if local is None:
            continue
        if met_open:                                   # remainders of a scrambled row
            if 'junho.' in ln:
                cut = ln.find('junho.') + len('junho.')
                buf.append(ln[:cut])
                if ln[cut:].strip():
                    robs.append(ln[cut:].strip())
                met_open = False
            elif re.match(r'^\d', ln):
                buf.append(ln)
            elif (len(ln.split()) <= 3 and ')' in ln and ln.count('(') <= ln.count(')')) \
                    or (len(ln.split()) == 1 and ln.islower() and ln.isalpha()):
                pextra.append(ln)
            else:
                robs.append(ln)
            continue
        if freq:
            robs.append(ln)
            continue
        fm = MON_FREQ.match(ln)
        if fm:
            freq = ln
            continue
        fm = re.match(rf'^({FREQ})\s+(.*)$', ln)       # 'Mensal <observações...>'
        if fm and 'metodologia' not in n:
            freq = fm.group(1)
            robs.append(fm.group(2))
            continue
        me = re.search(rf'\s({FREQ})$', ln)
        if me and 'metodologia' not in n:
            freq = me.group(1)
            buf.append(ln[:me.start()].strip())
            continue
        buf.append(ln)
    flush()
    return {'Observações': obs, 'Pontos': pontos, 'rows': rows}


def autocontrolo_from_lines(lines):
    seg = section(lines, 'Autocontrolo', ['Administrador', 'O Administrador', 'A Administradora',
                                          'Localização e caracterização'])
    if not seg and any(norm(l).startswith('programa de autocontrolo') for l in lines):
        seg = section(lines, 'Programa de autocontrolo',
                      ['Administrador', 'Localização e caracterização'])
    if not seg:
        return {}
    out = {'Observações': '', 'Periodicidade de reporte': '',
           'Equipamento de controlo': '', 'Notas de amostragem': '', 'rows': []}
    hdr = next((i for i, ln in enumerate(seg)
                if norm(ln).startswith('local de amostragem') or norm(ln) == 'local de'
                or (norm(ln).startswith('local') and 'parametro' in norm(ln))), None)
    rows_start = hdr + 1 if hdr is not None else \
        next((i for i, ln in enumerate(seg) if AUTO_LOCAL.match(ln)), len(seg))
    pre = seg[:hdr if hdr is not None else rows_start]
    for name, key, ends in (('observacoes', 'Observações',
                             ('periodicidade', 'descricao do equipamento', 'local')),
                            ('periodicidade de reporte', 'Periodicidade de reporte',
                             ('descricao do equipamento', 'observacoes', 'local')),
                            ('descricao do equipamento', 'Equipamento de controlo',
                             ('observacoes', 'periodicidade', 'local'))):
        i = next((i for i, ln in enumerate(pre) if norm(ln).startswith(name)), None)
        if i is None:
            continue
        first = re.sub(r'^[^:]*:\s*', '', pre[i]) if ':' in pre[i] else ''
        j = next((j for j in range(i + 1, len(pre))
                  if any(norm(pre[j]).startswith(e) for e in ends)), len(pre))
        out[key] = clean_text(first + ' ' + ' '.join(pre[i + 1:j]))
    buf, local, rows, pfreq = [], None, [], ''
    tail = []
    hdr_words = {'local', 'de', 'amostragem', 'parametro', 'metodo', 'analitico',
                 'frequencia', 'tipo'}

    def emit(freq, tipo):
        nonlocal local, pfreq
        before, met = _method_split(buf)
        rows.append({'Local de amostragem': local, 'Parâmetro': before,
                     'Método analítico': met,
                     'Frequência de amostragem': clean_text(freq),
                     'Tipo de amostragem': clean_text(tipo)})
        local, pfreq = None, ''

    for ln in seg[rows_start:]:
        n = norm(ln)
        if set(n.split()) <= hdr_words:
            continue                                   # wrapped header fragments
        if n.startswith('amostragem composta') or rows and n.startswith(('administrador', 'o administrador', 'a administradora')):
            if local is not None and pfreq:            # row waiting for its tipo cell
                emit(pfreq, '')
            tail.append(ln)
            local = None
            continue
        if tail:
            tail.append(ln)
            continue
        if pfreq:                                      # freq seen on its own line…
            if re.fullmatch(TIPO, ln.strip()):         # …tipo on the next line
                emit(pfreq, ln.strip())
                continue
            buf.append(pfreq)                          # false alarm — part of a cell
            pfreq = ''
        lm = AUTO_LOCAL.match(ln)
        if lm:
            if local is not None and buf:              # previous row never closed
                emit('', '')
            local, buf = lm.group(1), [ln[lm.end():].strip()]
            em = AUTO_END.search(buf[0])
            if em:                                      # single-line row (e.g. Caudal)
                buf = [buf[0][:em.start()].strip()]
                emit(em.group('freq'), em.group('tipo'))
            continue
        if local is None:
            continue
        em = AUTO_END.search(ln)
        if em:
            buf.append(ln[:em.start()].strip())
            emit(em.group('freq'), em.group('tipo'))
            continue
        if re.fullmatch(FREQ, ln.strip()):             # cell-per-line: bare freq line
            pfreq = ln.strip()
            continue
        buf.append(ln)
    if local is not None and pfreq:
        emit(pfreq, '')
    out['rows'] = rows
    out['Notas de amostragem'] = clean_text(' '.join(t for t in tail
                                                     if not norm(t).startswith(('administrador', 'o administrador', 'a administradora'))))
    return out


def _param_table_from_tables(rows, start_marker, header_req, cols, stop_markers):
    """Generic ruled-table reader for the monitorização/autocontrolo tables:
    marker -> header by NAME -> data rows, re-mapping columns on page-break
    header repeats (the TUA `_collect` pattern)."""
    m = next((i for i, r in enumerate(rows)
              if norm(' '.join(r)).startswith(norm(start_marker))), None)
    if m is None:
        return []
    out, ci = [], None
    for r in rows[m + 1:]:
        n = norm(' '.join(c for c in r if c))
        if any(n.startswith(norm(s)) for s in stop_markers):
            break
        if all(col_idx(r, h) is not None for h in header_req):
            ci = {name: col_idx(r, *terms) for name, terms in cols.items()}
            continue
        if ci is None:
            continue
        rec = {name: clean_text(r[i]) if i is not None and i < len(r) else ''
               for name, i in ci.items()}
        first = next(iter(cols))
        if rec.get(first) and rec.get('Parâmetro'):
            out.append(rec)
    return out


def autocontrolo_from_tables(rows):
    return _param_table_from_tables(
        rows, 'programa de autocontrolo a implementar',
        ['local de amostragem', 'parametro', 'frequencia'],
        {'Local de amostragem': ('local de amostragem',), 'Parâmetro': ('parametro',),
         'Método analítico': ('metodo analitico',),
         'Frequência de amostragem': ('frequencia de amostragem',),
         'Tipo de amostragem': ('tipo de amostragem',)},
        ['amostragem composta', 'administrador', 'localizacao e caracterizacao'])


def monitorizacao_from_tables(rows):
    return _param_table_from_tables(
        rows, 'programa de monitorizacao do meio recetor',
        ['local', 'parametro', 'frequencia'],
        {'Local': ('local',), 'Parâmetro': ('parametro',),
         'Método analítico': ('metodo analitico',),
         'Frequência': ('frequencia',), 'Observações': ('observacoes',)},
        ['autocontrolo', 'programa de autocontrolo'])


def labels_from_tables(rows, labels):
    """Label/value pairs from 2-column ruled rows — the cleanest source when the
    ruling is detected, since cell boundaries already separate label from value."""
    matcher = _prep_labels(labels)
    out = {}
    for r in rows:
        cells = [c for c in r if c is not None]
        if len(cells) < 2 or not cells[0]:
            continue
        m = match_label(cells[0], matcher)
        if m is None:
            continue
        idx, extra = m
        name = labels[idx][0]
        val = clean_text(' '.join([extra] + [c for c in cells[1:] if c]))
        if val and not out.get(name):
            out[name] = val
    return out


def entidade(lines):
    for i, ln in enumerate(lines):
        m = re.match(r'^(?:[OA]\s+)?Administrador(?:a)?\s+Regional\s+da\s+(ARH\s+.+)$', ln)
        if m:
            signer = ''
            for nxt in lines[i + 1:i + 5]:             # skip '(Ao abrigo da …)' lines
                s = nxt.strip()
                if s.startswith('(') or s.endswith(')'):
                    continue
                if 1 <= len(s.split()) <= 6 and not s[:1].islower():
                    signer = s
                break
            return {'Entidade Licenciadora': f'APA — {clean_text(m.group(1))}',
                    'Assinado por': clean_text(signer)}
    return {}


# ----------------------------------------------------------------- top level
def extract(pdf_path):
    pages, table_rows = load_pdf(pdf_path)
    return extract_from_parts(pages, table_rows, os.path.basename(pdf_path))


def extract_from_parts(pages, table_rows, file_name):
    """The whole extraction from (text pages, ruled-table rows). Ruled tables are
    preferred for each section; the text-layer parser fills anything the ruling
    missed — so no single detection failure can blank a section."""
    lines = doc_lines(pages)
    warnings = []

    dg = dados_gerais(lines)
    trat = tratamento(lines)
    rej = rejeicao(lines)
    if table_rows:                                   # overlay: table cells win
        for d, labels in ((dg, HEADER_LABELS + IDENT_LABELS), (trat, TRAT_LABELS), (rej, REJ_LABELS)):
            for k, v in labels_from_tables(table_rows, labels).items():
                if v and len(v) >= len(d.get(k, '')):
                    d[k] = v
        for k in ('Data de Início', 'Data de Validade'):
            if dg.get(k):
                dg[k] = date_norm(dg[k])
        for d, k in ((dg, 'Obrigação de correcção de Dados de Perfil'),
                     (rej, 'Valorização ou reutilização')):
            if d.get(k):
                d[k] = checkbox(d[k])
        if dg.get('Nº Licença'):
            m = LIC_RE.search(dg['Nº Licença'].replace(' ', ''))
            if m:
                dg['Nº Licença'] = m.group(0)

    conds_t = vle_from_tables(table_rows) if table_rows else []
    text_warn = []
    conds_l, conds_obs = vle_from_lines(lines, text_warn)
    if len(conds_t) >= len(conds_l):                 # ruled tables win ties (cleaner cells)
        conds = conds_t or conds_l
        if conds is conds_l:
            warnings += text_warn
    else:                                            # text reader saw more rows — keep them
        conds = conds_l
        warnings += text_warn
        if conds_t:
            warnings.append(f'ruled-table reader found only {len(conds_t)} of '
                            f'{len(conds_l)} discharge rows — used the text reading')

    legis, aval, aval_texto = legends(lines)
    for c in conds:
        lt, at = resolve(c.get('Legislação aplicável'), legis, aval, aval_texto)
        c['Legislação aplicável (texto)'] = lt
        c['Avaliação da conformidade (texto)'] = at

    mon = monitorizacao_from_lines(lines)
    mrows = monitorizacao_from_tables(table_rows) if table_rows else []
    if mrows and len(mrows) >= len(mon.get('rows', [])):
        mon['rows'] = mrows

    auto = autocontrolo_from_lines(lines)
    arows = autocontrolo_from_tables(table_rows) if table_rows else []
    if arows and len(arows) >= len(auto.get('rows', [])):
        auto['rows'] = arows

    data = {
        'file_name': file_name,
        'dados_gerais': dg,
        'tratamento': trat,
        'rejeicao': rej,
        'condicoes_texto': condicoes_texto(lines),
        'condicoes_descarga': conds,
        'condicoes_observacoes': conds_obs,
        'legislacao': legis,
        'avaliacao_conformidade': aval,
        'avaliacao_texto': aval_texto,
        'monitorizacao': mon,
        'autocontrolo': auto,
        'entidade': entidade(lines),
        '_warnings': warnings,
    }
    data['_warnings'] += _warnings(data)
    return data


def _warnings(d):
    w = []
    if not d['dados_gerais'].get('Nº Licença'):
        w.append('Nº Licença not found')
    if not d['dados_gerais'].get('Data de Validade'):
        w.append('Data de Validade not found')
    if not d['condicoes_descarga']:
        w.append('No discharge conditions (VLE table) found')
    if not d['autocontrolo'].get('rows'):
        w.append('No autocontrolo rows found')
    unit = re.compile(r'\((?:mg/L|Escala|ºC|m3)')
    for label, items in (('condition', d['condicoes_descarga']),
                         ('autocontrolo', d['autocontrolo'].get('rows', []))):
        garbled = [it.get('Parâmetro', '') for it in items
                   if len(unit.findall(it.get('Parâmetro', ''))) > 1]
        if garbled:
            w.append(f"{label} parameter(s) look garbled (page-break fragmentation): {garbled}")
    return w


def _status(validade):
    try:
        dt = datetime.strptime(validade, '%d-%m-%Y').date()
        return 'Em vigor' if dt >= datetime.now().date() else 'Caducada'
    except Exception:
        return ''


def lic_key(d):
    """Stable key linking every sheet to its licence (Nº Licença, file_name fallback)."""
    return d['dados_gerais'].get('Nº Licença', '') or d['file_name']


# ------------------------------------------------------------- master sheets
def license_row(d):
    g, t, r = d['dados_gerais'], d['tratamento'], d['rejeicao']
    e = d['entidade']
    return {
        'file_name': d['file_name'], 'TUA/LURH': 'LURH',
        'Nº Licença': g.get('Nº Licença', ''), 'Nº Processo': g.get('Nº Processo', ''),
        'Estabelecimento': t.get('Designação', ''), 'Código APA': g.get('Código APA', ''),
        'Requerente': g.get('Requerente', ''), 'NIF': g.get('NIF', ''),
        'Data de Início': g.get('Data de Início', ''),
        'Data de Validade': g.get('Data de Validade', ''),
        'Em Vigor/Caducada': _status(g.get('Data de Validade', '')),
        'Entidade Licenciadora': e.get('Entidade Licenciadora', ''),
        'Região Hidrográfica': r.get('Região Hidrográfica', ''),
        'Bacia Hidrográfica': r.get('Bacia Hidrográfica', ''),
        'Sub-Bacia Hidrográfica': r.get('Sub-Bacia Hidrográfica', ''),
        'Tipo de massa de água': r.get('Tipo de massa de água', ''),
        'Massa de Água': r.get('Massa de água', ''),
        'Classificação da Massa de Água': r.get('Classificação da massa de água', ''),
        'Meio Recetor': r.get('Meio Recetor', ''),
        'Denominação do meio recetor': r.get('Denominação do meio recetor', ''),
        'Margem': r.get('Margem', ''), 'Sistema de Descarga': r.get('Sistema de Descarga', ''),
        'Longitude (descarga)': r.get('Longitude', ''), 'Latitude (descarga)': r.get('Latitude', ''),
        'Longitude (ETAR)': t.get('Longitude', ''), 'Latitude (ETAR)': t.get('Latitude', ''),
        'Nut III – Concelho – Freguesia': t.get('Nut III – Concelho – Freguesia', ''),
        'Nível de tratamento': t.get('Nível de tratamento', ''),
        'Tipo de tratamento': t.get('Tipo de tratamento', ''),
        'Caudal Máximo descarga': t.get('Caudal Máximo descarga', ''),
        'Caudal Médio descarga': t.get('Caudal Médio descarga', ''),
        'Ano de arranque': t.get('Ano de arranque', ''),
        'População servida (e.p.)': t.get('População servida (e.p.)', ''),
        'Ano horizonte de projeto': t.get('Ano horizonte de projeto', ''),
        'População servida no horizonte': t.get('População servida no horizonte', ''),
        'Origem das águas residuais': r.get('Origem das águas residuais', ''),
        'Volume mensal afluente (m3)': r.get('Volume mensal afluente (m3)', ''),
        'Afluente CBO5 (mg/L)': r.get('Afluente CBO5 (mg/L)', ''),
        'Afluente CQO (mg/L)': r.get('Afluente CQO (mg/L)', ''),
        'Afluente N (mg/L)': r.get('Afluente N (mg/L)', ''),
        'Afluente P (mg/L)': r.get('Afluente P (mg/L)', ''),
        'Valorização ou reutilização': r.get('Valorização ou reutilização', ''),
        'Caudal Reutilizado': r.get('Caudal Reutilizado', ''),
        'Finalidades Efluente': r.get('Finalidades Efluente', ''),
    }


def condition_rows(d):
    """One row per VLE parameter, with the matching Saída autocontrolo frequency
    merged in (the same join the TUA master does)."""
    lic = lic_key(d)
    saida = {norm(a['Parâmetro']): a for a in d['autocontrolo'].get('rows', [])
             if norm(a.get('Local de amostragem', '')).startswith('sai')}
    out = []
    for c in d['condicoes_descarga']:
        ac = saida.get(norm(c.get('Parâmetro', '')), {})
        out.append({'Nº Licença': lic, 'Regime': c.get('Regime', 'Normal'),
                    'Parâmetro': c.get('Parâmetro', ''),
                    'VLE (% mín. remoção)': c.get('VLE (% mín. remoção)', ''),
                    'VLE': c.get('VLE', ''),
                    'Legislação aplicável': c.get('Legislação aplicável (texto)', ''),
                    'Avaliação da Conformidade Legal': c.get('Avaliação da conformidade (texto)', ''),
                    'Frequência de amostragem': ac.get('Frequência de amostragem', ''),
                    'Tipo de amostragem': ac.get('Tipo de amostragem', '')})
    return out


def autocontrolo_rows(d):
    lic = lic_key(d)
    return [{'Nº Licença': lic, 'Local de amostragem': a.get('Local de amostragem', ''),
             'Parâmetro': a.get('Parâmetro', ''),
             'Método analítico': a.get('Método analítico', ''),
             'Frequência de amostragem': a.get('Frequência de amostragem', ''),
             'Tipo de amostragem': a.get('Tipo de amostragem', '')}
            for a in d['autocontrolo'].get('rows', [])]


def meio_recetor_rows(d):
    lic = lic_key(d)
    return [{'Nº Licença': lic, 'Local': m.get('Local', ''),
             'Parâmetro': m.get('Parâmetro', ''),
             'Método analítico': m.get('Método analítico', ''),
             'Frequência': m.get('Frequência', ''),
             'Observações': m.get('Observações', '')}
            for m in d['monitorizacao'].get('rows', [])]


def legend_rows(d):
    lic = lic_key(d)
    legis = [{'Nº Licença': lic, 'Código': k, 'Legislação aplicável': v}
             for k, v in sorted(d['legislacao'].items())]
    aval = [{'Nº Licença': lic, 'Código': k, 'Avaliação da conformidade': v}
            for k, v in sorted(d['avaliacao_conformidade'].items())]
    if not aval and d.get('avaliacao_texto'):
        aval = [{'Nº Licença': lic, 'Código': '', 'Avaliação da conformidade': d['avaliacao_texto']}]
    return legis, aval


MASTER_SHEETS = ['Licenses', 'Conditions', 'Autocontrolo', 'MeioRecetor', 'Legislacao', 'Avaliacao']


def build_sheets(d):
    legis, aval = legend_rows(d)
    return {'Licenses': [license_row(d)], 'Conditions': condition_rows(d),
            'Autocontrolo': autocontrolo_rows(d), 'MeioRecetor': meio_recetor_rows(d),
            'Legislacao': legis, 'Avaliacao': aval}


# --------------------------------------------------- per-file workbook (human view)
_HDR_FILL = PatternFill('solid', fgColor='1F4E78')
_HDR_FONT = Font(bold=True, color='FFFFFF')
_SEC_FILL = PatternFill('solid', fgColor='D9E1F2')
_SEC_FONT = Font(bold=True, color='1F4E78')
_TITLE_FONT = Font(bold=True, size=14, color='1F4E78')
_WRAP = Alignment(wrap_text=True, vertical='top')
_TOP = Alignment(vertical='top')


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _table(ws, headers, rows, widths, wrap_cols=()):
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font = _HDR_FILL, _HDR_FONT
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 28
    for r in rows:
        ws.append([v if v not in (None, '') else '' for v in r])
    _widths(ws, widths)
    wrap = set(wrap_cols)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = _WRAP if cell.column in wrap else _TOP
    ws.freeze_panes = 'A2'


def _summary(ws, d):
    g, t, r, e = d['dados_gerais'], d['tratamento'], d['rejeicao'], d['entidade']
    estado = _status(g.get('Data de Validade', ''))
    blocks = [
        ('Identificação', [
            ('Nº Licença', g.get('Nº Licença')), ('Nº Processo', g.get('Nº Processo')),
            ('Código APA', g.get('Código APA')), ('Requerente', g.get('Requerente')),
            ('NIF', g.get('NIF')),
        ]),
        ('Validade da licença', [
            ('Data de Início', g.get('Data de Início')),
            ('Data de Validade', g.get('Data de Validade')),
            ('Estado', estado),
            ('Entidade Licenciadora', e.get('Entidade Licenciadora')),
        ]),
        ('Localização da descarga', [
            ('Designação da rejeição', r.get('Designação da rejeição')),
            ('Meio Recetor', r.get('Meio Recetor')),
            ('Denominação do meio recetor', r.get('Denominação do meio recetor')),
            ('Massa de Água', r.get('Massa de água')),
            ('Classificação da Massa de Água', r.get('Classificação da massa de água')),
            ('Região Hidrográfica', r.get('Região Hidrográfica')),
            ('Longitude', r.get('Longitude')), ('Latitude', r.get('Latitude')),
        ]),
        ('Estação e tratamento', [
            ('Designação', t.get('Designação')),
            ('Nível de tratamento', t.get('Nível de tratamento')),
            ('Tipo de tratamento', t.get('Tipo de tratamento')),
            ('Caudal Máximo descarga', t.get('Caudal Máximo descarga')),
            ('Caudal Médio descarga', t.get('Caudal Médio descarga')),
            ('População servida (e.p.)', t.get('População servida (e.p.)')),
            ('Ano horizonte de projeto', t.get('Ano horizonte de projeto')),
            ('População servida no horizonte', t.get('População servida no horizonte')),
        ]),
    ]
    ws['A1'] = t.get('Designação') or d['file_name']
    ws['A1'].font = _TITLE_FONT
    ws['A2'] = f'Licença de Utilização dos Recursos Hídricos — {estado}' if estado \
               else 'Licença de Utilização dos Recursos Hídricos'
    ws['A2'].font = Font(italic=True, color='595959')
    ws.merge_cells('A1:B1')
    ws.merge_cells('A2:B2')
    row = 4
    for title, kvs in blocks:
        row = _kv_block(ws, row, title, kvs) + 1
    for msg in (d.get('_warnings') or []):
        ws.cell(row, 1, '⚠ Verificar').font = Font(bold=True, color='C00000')
        ws.cell(row, 2, msg).alignment = _WRAP
        row += 1
    _widths(ws, [32, 72])


def _kv_block(ws, row, title, kvs):
    """A titled label/value block; returns the next free row."""
    ws.cell(row, 1, title).font = _SEC_FONT
    ws.cell(row, 1).fill = ws.cell(row, 2).fill = _SEC_FILL
    row += 1
    for label, val in kvs:
        ws.cell(row, 1, label).font = Font(bold=True)
        ws.cell(row, 2, val if val not in (None, '') else '—').alignment = _WRAP
        row += 1
    return row


def _fields_sheet(ws, d):
    """Every extracted label/value field, grouped and ordered like the LURH
    document itself — so nothing that was read is missing from the Excel."""
    g, t, r, e = d['dados_gerais'], d['tratamento'], d['rejeicao'], d['entidade']
    estado = _status(g.get('Data de Validade', ''))
    row = 1
    row = _kv_block(ws, row, 'Licença de Utilização dos Recursos Hídricos',
                    [(name, g.get(name)) for name, _ in HEADER_LABELS] +
                    [('Estado', estado)]) + 1
    row = _kv_block(ws, row, 'Identificação',
                    [(name, g.get(name)) for name, _ in IDENT_LABELS]) + 1
    row = _kv_block(ws, row, 'Caracterização do sistema de tratamento',
                    [(name, t.get(name)) for name, _ in TRAT_LABELS]) + 1
    row = _kv_block(ws, row, 'Caracterização da rejeição',
                    [(name, r.get(name)) for name, _ in REJ_LABELS]) + 1
    row = _kv_block(ws, row, 'Características do Afluente Bruto',
                    [('Volume mensal afluente (m3)', r.get('Volume mensal afluente (m3)')),
                     ('Volume mensal afluente (tipo)', r.get('Volume mensal afluente (tipo)')),
                     ('Afluente CBO5 (mg/L)', r.get('Afluente CBO5 (mg/L)')),
                     ('Afluente CQO (mg/L)', r.get('Afluente CQO (mg/L)')),
                     ('Afluente N (mg/L)', r.get('Afluente N (mg/L)')),
                     ('Afluente P (mg/L)', r.get('Afluente P (mg/L)'))]) + 1
    _kv_block(ws, row, 'Entidade Licenciadora',
              [('Entidade Licenciadora', e.get('Entidade Licenciadora')),
               ('Assinado por', e.get('Assinado por'))])
    _widths(ws, [40, 80])


def _notes_below(ws, pairs):
    """Label/value notes appended under a table (blank row first)."""
    pairs = [(l, v) for l, v in pairs if v not in (None, '')]
    if not pairs:
        return
    ws.append([])
    for label, val in pairs:
        ws.append([label, val])
        c = ws.cell(ws.max_row, 1)
        c.font = Font(bold=True, color='1F4E78')
        c.alignment = _TOP
        ws.cell(ws.max_row, 2).alignment = _WRAP


def write_outputs(d, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(d['file_name'])[0]
    json_path = os.path.join(out_dir, base + '_extracted.json')
    xlsx_path = os.path.join(out_dir, base + '_extracted.xlsx')

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

    wb = Workbook()
    _summary(wb.active, d)
    wb.active.title = 'Resumo'

    _fields_sheet(wb.create_sheet('Dados da Licença'), d)

    cond = condition_rows(d)
    ws = wb.create_sheet('Condições de Descarga')
    _table(ws,
           ['Regime', 'Parâmetro', 'VLE (% mín. remoção)', 'VLE',
            'Frequência', 'Tipo de amostragem', 'Legislação aplicável', 'Avaliação da conformidade'],
           [[c.get('Regime'), c.get('Parâmetro'), c.get('VLE (% mín. remoção)'), c.get('VLE'),
             c.get('Frequência de amostragem'), c.get('Tipo de amostragem'),
             c.get('Legislação aplicável'), c.get('Avaliação da Conformidade Legal')] for c in cond],
           [16, 30, 14, 12, 16, 16, 46, 46], wrap_cols=(2, 7, 8))
    _notes_below(ws, [(f'Observações — {reg}', txt)
                      for reg, txt in (d.get('condicoes_observacoes') or {}).items()])

    auto = d['autocontrolo']
    ws = wb.create_sheet('Autocontrolo')
    _table(ws,
           ['Local de amostragem', 'Parâmetro', 'Frequência', 'Tipo de amostragem', 'Método analítico'],
           [[a.get('Local de amostragem'), a.get('Parâmetro'), a.get('Frequência de amostragem'),
             a.get('Tipo de amostragem'), a.get('Método analítico')]
            for a in auto.get('rows', [])],
           [20, 30, 16, 18, 60], wrap_cols=(2, 5))
    _notes_below(ws, [('Observações', auto.get('Observações')),
                      ('Periodicidade de reporte', auto.get('Periodicidade de reporte')),
                      ('Equipamento de controlo', auto.get('Equipamento de controlo')),
                      ('Notas de amostragem', auto.get('Notas de amostragem'))])

    mon = d['monitorizacao']
    if mon.get('rows') or mon.get('Pontos') or mon.get('Observações'):
        ws = wb.create_sheet('Meio Recetor')
        _table(ws,
               ['Local', 'Parâmetro', 'Frequência', 'Observações', 'Método analítico'],
               [[m.get('Local'), m.get('Parâmetro'), m.get('Frequência'),
                 m.get('Observações'), m.get('Método analítico')]
                for m in mon.get('rows', [])],
               [14, 30, 14, 30, 60], wrap_cols=(2, 4, 5))
        _notes_below(ws, [(f'Ponto {p}', txt) for p, txt in (mon.get('Pontos') or {}).items()] +
                         [('Observações da secção', mon.get('Observações'))])

    ct = d['condicoes_texto']
    _table(wb.create_sheet('Condições (texto)'),
           ['Secção', 'Nº', 'Condição'],
           [[sec, i + 1, txt] for sec, items in
            (('Gerais', ct['Gerais']), ('Específicas', ct['Específicas']), ('Outras', ct['Outras']))
            for i, txt in enumerate(items)],
           [14, 6, 120], wrap_cols=(3,))

    legis, aval = legend_rows(d)
    refs = [['Legislação', x['Código'], x['Legislação aplicável']] for x in legis] + \
           [['Avaliação', x['Código'], x['Avaliação da conformidade']] for x in aval]
    _table(wb.create_sheet('Referências'), ['Tipo', 'Código', 'Texto'], refs,
           [14, 10, 110], wrap_cols=(3,))

    wb.save(xlsx_path)
    return json_path, xlsx_path


def update_master(d, master_path):
    """Upsert this licence into the normalized multi-sheet master workbook, all
    linked on Nº Licença. Re-extracting a licence REPLACES its rows in every
    sheet. Robust to the master being open in Excel (diverts to .PENDING.xlsx)."""
    new_sheets = build_sheets(d)
    lic = str(lic_key(d))

    existing, read_failed = {}, False
    if os.path.isfile(master_path):
        try:
            existing = pd.read_excel(master_path, sheet_name=None)
        except Exception:
            read_failed = True
    result = {}
    for name in MASTER_SHEETS:
        new_df = pd.DataFrame(new_sheets[name])
        old = existing.get(name, pd.DataFrame())
        if not old.empty and 'Nº Licença' in old.columns:
            old = old[old['Nº Licença'].astype(str) != lic]
        result[name] = pd.concat([old, new_df], ignore_index=True)

    pending = os.path.splitext(master_path)[0] + '.PENDING.xlsx'
    target, deferred = (pending, True) if read_failed else (master_path, False)
    if not deferred and os.path.isfile(master_path):
        try:
            shutil.copy2(master_path, os.path.splitext(master_path)[0] + '.bak.xlsx')
        except Exception:
            pass
    try:
        with pd.ExcelWriter(target, engine='openpyxl') as xl:
            for name in MASTER_SHEETS:
                result[name].to_excel(xl, sheet_name=name, index=False)
    except PermissionError:
        target, deferred = pending, True
        with pd.ExcelWriter(target, engine='openpyxl') as xl:
            for name in MASTER_SHEETS:
                result[name].to_excel(xl, sheet_name=name, index=False)
    return target, {n: len(result[n]) for n in MASTER_SHEETS}, deferred


def export_master_csv(master_path, csv_dir=None):
    """Each master sheet as a UTF-8 CSV — the stable Power BI source."""
    if not os.path.isfile(master_path):
        return None
    csv_dir = csv_dir or os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(master_path)))), 'data', 'powerbi'), 'LURH')
    os.makedirs(csv_dir, exist_ok=True)
    try:
        sheets = pd.read_excel(master_path, sheet_name=None)
    except Exception:
        return None
    try:
        import enrich_lurh                              # optional, like TUA's enrich
        sheets, _ = enrich_lurh.enrich_all(sheets)
    except Exception:
        pass
    for name, df in sheets.items():
        df.to_csv(os.path.join(csv_dir, name + '.csv'), index=False, encoding='utf-8-sig')
    return csv_dir


def resolve_pdf_path(argv):
    if len(argv) > 1:
        return argv[1]
    env = os.getenv('EPAL_LURH_PDF_FILE')
    if env:
        return env
    here = os.path.dirname(os.path.abspath(__file__))
    pdf_dir = os.path.join(here, 'PDF files')
    if os.path.isdir(pdf_dir):
        for f in sorted(os.listdir(pdf_dir)):
            if f.lower().endswith('.pdf'):
                return os.path.join(pdf_dir, f)
    return None


def main(argv):
    pdf = resolve_pdf_path(argv)
    if not pdf or not os.path.isfile(pdf):
        print(f"ERROR: LURH PDF not found. Pass a path: python extract_lurh.py <file.pdf>  (got: {pdf})")
        return 2
    out_dir = os.getenv('EPAL_LURH_OUTPUT_DIR', os.path.dirname(os.path.abspath(__file__)))
    data = extract(pdf)
    jp, xp = write_outputs(data, out_dir)
    print(f"Extracted {len(data['condicoes_descarga'])} discharge parameters, "
          f"{len(data['autocontrolo'].get('rows', []))} autocontrolo rows and "
          f"{len(data['monitorizacao'].get('rows', []))} meio-recetor rows from {data['file_name']}")
    if data['_warnings']:
        print("WARNINGS:", '; '.join(data['_warnings']))
    print("JSON :", jp)
    print("Excel:", xp)
    master_path = os.getenv('EPAL_LURH_MASTER', os.path.join(out_dir, 'data', 'reference', 'master_lurh.xlsx'))
    mp, counts, deferred = update_master(data, master_path)
    summary = ', '.join(f"{n}={counts[n]}" for n in MASTER_SHEETS)
    if deferred:
        print(f"WARNING: '{master_path}' is open/locked in Excel — the master was NOT updated.")
        print(f"         Saved to '{mp}' instead. Close the master and re-run to consolidate.")
        return 3
    print(f"Master: {mp}  ({summary})")
    csv_dir = export_master_csv(master_path)
    if csv_dir:
        print("Power BI CSVs:", csv_dir)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
