"""Canonical criteria model shared by the TUA and LURH criteria workbooks.

Both `src/tua/build_criteria_table.py` and
`src/lurh/build_criteria_table.py` import this module, so the two
workbooks are guaranteed to carry the SAME sheet names, the SAME column
headers in the SAME order, and the same ID conventions. That is what makes
them appendable in Power BI (Append Queries) without any per-source mapping.

Star schema (7 sheets, identical in both files)
-----------------------------------------------
  Leia-me              documentation
  Catalogo_Criterios   IdCriterio -> parâmetro + VLE + the 9 criteria flags
  Textos_Legais        IdTexto    -> legislação + avaliação + interpretação
  Licencas_Criterios   fact table: licença x parâmetro -> IdCriterio + IdTexto
  Dim_Licencas         IdLicenca  -> estabelecimento, empresa, ficheiro fonte
  Dim_Parametros       nome no documento -> nome no catálogo
  Vista_Plana          everything joined back into one flat table (regenerated)

Key conventions
---------------
  IdLicenca   the Nº TUA for TUA, the Nº Licença for LURH
  IdCriterio  '<FONTE>_<PARAM>_<n>'  e.g. TUA_CQO_1, LURH_CQO_1
  IdTexto     '<FONTE>_TX<nn>'       e.g. TUA_TX01, LURH_TX07
  Fonte       'TUA' or 'LURH' — present on every table, so an appended model
              can always be sliced back apart
  Regime      'Normal' for every TUA row (the concept does not exist in a TUA);
              'Normal' / 'Ano de arranque' / … for LURH

The prefix on the IDs is what keeps the two files collision-free when appended.
"""
import os
import re
import unicodedata
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------- schema
FONTE_TUA, FONTE_LURH = 'TUA', 'LURH'

FLAG_COLS = [
    'VLE (flag)', '% mín remoção (flag)', 'Gama de valores (intervalo)',
    '≤ 100% VLE (dobro)', '≤ 150% VLE', '≤ uma ordem de grandeza do VLE',
    'média mensal ≤ VLE', 'média anual ≤ VLE', 'Quadro III DL 152/97 (borlas)',
]

CATALOGO_COLS = (['IdCriterio', 'Fonte', 'Empresa (catálogo)', 'Parâmetro (catálogo)', 'VLE']
                 + FLAG_COLS + ['Origem dos critérios', 'Nº de linhas que o usam'])
TEXTOS_COLS = ['IdTexto', 'Fonte', 'Legislação aplicável (texto)',
               'Texto da avaliação de conformidade (lei)', 'Interpretação → critérios',
               'Nº de linhas que o usam']
FACTOS_COLS = ['IdLicenca', 'Fonte', 'Regime', 'Parâmetro (documento)', 'VLE',
               'VLE (% mín. remoção)', 'Nota/Código', 'IdCriterio', 'IdTexto']
LICENCAS_COLS = ['IdLicenca', 'Fonte', 'Estabelecimento', 'Empresa', 'Ficheiro fonte']
PARAMETROS_COLS = ['Parâmetro (documento)', 'Parâmetro (catálogo)', 'Fonte']
VISTA_COLS = (['IdLicenca', 'Fonte', 'Estabelecimento', 'Empresa', 'Ficheiro fonte',
               'Regime', 'Parâmetro (documento)', 'Parâmetro (catálogo)', 'VLE',
               'VLE (% mín. remoção)', 'Nota/Código', 'Interpretação → critérios',
               'Texto da avaliação de conformidade (lei)', 'Legislação aplicável (texto)']
              + FLAG_COLS + ['Empresa (catálogo)', 'IdCriterio', 'IdTexto',
                             'Origem dos critérios'])

SHEETS = ['Leia-me', 'Catalogo_Criterios', 'Textos_Legais', 'Licencas_Criterios',
          'Dim_Licencas', 'Dim_Parametros', 'Vista_Plana']

# provenance values for 'Origem dos critérios'
ORIG_MANUAL = 'Manual (validado)'
ORIG_GLOBAL = 'Herdado de CriteriosConformidade_GLOBAL'
ORIG_AUTO = 'AUTO — rever'


# ------------------------------------------------------ shared vocabulary
def norm(s):
    """Accent/case/whitespace-folded comparison key."""
    if s is None:
        return ''
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().lower()


# One catalogue vocabulary for both document families. Matched on the folded
# parameter name, longest pattern first, so 'azoto amoniacal' never falls into
# the 'azoto total' bucket.
PARAM_PATTERNS = [
    ('carencia bioquimica', 'CBO5'),
    ('cbo5', 'CBO5'),
    ('carencia quimica', 'CQO'),
    ('cqo', 'CQO'),
    ('solidos suspensos', 'SST'),
    ('particulas solidas em suspensao', 'SST'),
    ('sst', 'SST'),
    ('azoto amoniacal', 'Azoto amoniacal'),
    ('azoto total', 'Nt'),
    ('fosforo total', 'Pt'),
    ('escherichia coli', 'E. coli'),
    ('e. coli', 'E. coli'),
    ('coliformes fecais', 'Coliformes fecais'),
    ('oleos e gorduras', 'Óleo e gorduras'),
    ('fenois', 'Fenóis'),
    ('caudal', 'Caudal'),
    ('ph', 'pH'),
]


def param_catalogo(name):
    """Document parameter name -> catalogue name (shared by TUA and LURH)."""
    n = norm(name)
    for pat, cat in PARAM_PATTERNS:
        if pat in n:
            return cat
    return (str(name or '').strip() or 'Desconhecido')


def slug(name):
    """'Azoto amoniacal' -> 'AZOTO_AMONIACAL' (safe inside an ID)."""
    s = unicodedata.normalize('NFKD', str(name or '')).encode('ascii', 'ignore').decode()
    s = re.sub(r'[^A-Za-z0-9]+', '_', s).strip('_').upper()
    return s or 'PARAM'


def vle_sort_key(v):
    """Numeric VLEs sort numerically; ranges and text sort after, alphabetically."""
    s = str(v or '').strip().replace(',', '.')
    m = re.fullmatch(r'\d+(?:\.\d+)?', s)
    return (0, float(m.group()), '') if m else (1, 0.0, s)


# ------------------------------------------------------------ normalise
def normalise(rows, fonte):
    """Flat records -> (catalogo, textos, factos, licencas, parametros).

    Each input record is a dict with the keys:
      IdLicenca, Estabelecimento, Empresa, Ficheiro fonte, Regime,
      Parâmetro (documento), Parâmetro (catálogo), VLE, VLE (% mín. remoção),
      Nota/Código, Interpretação → critérios,
      Texto da avaliação de conformidade (lei), Legislação aplicável (texto),
      Empresa (catálogo), Origem dos critérios, flags (9-tuple)

    IdCriterio is keyed on (Parâmetro catálogo, VLE, flags) — VLE stays in the
    key so VLE-specific interpretations remain visible as separate rules.
    The legal texts are a SEPARATE dimension because the same rule appears with
    different wording across licences; folding them in would lose information.
    """
    cat_key, per_param = {}, {}
    for r in sorted(rows, key=lambda r: (str(r['Parâmetro (catálogo)'] or ''),
                                         vle_sort_key(r['VLE']), r['flags'])):
        k = (r['Parâmetro (catálogo)'], r['VLE'], r['flags'])
        if k in cat_key:
            continue
        p = slug(r['Parâmetro (catálogo)'])
        per_param[p] = per_param.get(p, 0) + 1
        cat_key[k] = f'{fonte}_{p}_{per_param[p]}'

    txt_key, seq = {}, 0
    for r in sorted(rows, key=lambda r: (str(r['Legislação aplicável (texto)'] or ''),
                                         str(r['Texto da avaliação de conformidade (lei)'] or ''),
                                         str(r['Interpretação → critérios'] or ''))):
        k = (r['Interpretação → critérios'],
             r['Texto da avaliação de conformidade (lei)'],
             r['Legislação aplicável (texto)'])
        if k not in txt_key:
            seq += 1
            txt_key[k] = f'{fonte}_TX{seq:02d}'

    def ck(r):
        return cat_key[(r['Parâmetro (catálogo)'], r['VLE'], r['flags'])]

    def tk(r):
        return txt_key[(r['Interpretação → critérios'],
                        r['Texto da avaliação de conformidade (lei)'],
                        r['Legislação aplicável (texto)'])]

    cat_use, txt_use, cat_emp, cat_orig = {}, {}, {}, {}
    for r in rows:
        cat_use[ck(r)] = cat_use.get(ck(r), 0) + 1
        txt_use[tk(r)] = txt_use.get(tk(r), 0) + 1
        cat_emp.setdefault(ck(r), set()).add(r.get('Empresa (catálogo)'))
        cat_orig.setdefault(ck(r), set()).add(r.get('Origem dos critérios'))

    catalogo = []
    for (param, vle, flags), cid in sorted(
            cat_key.items(), key=lambda kv: (kv[1].rsplit('_', 1)[0],
                                             int(kv[1].rsplit('_', 1)[1]))):
        emp = sorted(e for e in cat_emp[cid] if e)
        orig = sorted(o for o in cat_orig[cid] if o)
        catalogo.append([cid, fonte, ' ; '.join(emp) or None, param, vle] + list(flags)
                        + [' ; '.join(orig) or None, cat_use[cid]])

    textos = [[tid, fonte, legis, aval, interp, txt_use[tid]]
              for (interp, aval, legis), tid in sorted(txt_key.items(), key=lambda kv: kv[1])]

    factos = [[r['IdLicenca'], fonte, r['Regime'], r['Parâmetro (documento)'], r['VLE'],
               r['VLE (% mín. remoção)'], r['Nota/Código'], ck(r), tk(r)] for r in rows]

    lic = {}
    for r in rows:
        lic.setdefault(r['IdLicenca'], [r['IdLicenca'], fonte, r['Estabelecimento'],
                                        r['Empresa'], r['Ficheiro fonte']])
    licencas = sorted(lic.values(), key=lambda x: str(x[2] or ''))

    par = {}
    for r in rows:
        k = r['Parâmetro (documento)']
        par.setdefault(k, [k, r['Parâmetro (catálogo)'], fonte])
        if par[k][1] != r['Parâmetro (catálogo)']:
            raise SystemExit(f"'{k}' maps to two catalogue names — Dim_Parametros "
                             f"would not be a function")
    parametros = sorted(par.values(), key=lambda x: (str(x[1] or ''), str(x[0] or '')))

    return catalogo, textos, factos, licencas, parametros


def flatten(catalogo, textos, factos, licencas, parametros):
    """Rebuild the flat view from the normalised tables (this is Vista_Plana)."""
    cat = {r[0]: r for r in catalogo}
    txt = {r[0]: r for r in textos}
    lic = {r[0]: r for r in licencas}
    par = {r[0]: r for r in parametros}
    out = []
    for idlic, fonte, regime, pdoc, vle, pct, nota, cid, tid in factos:
        L, C, T = lic[idlic], cat[cid], txt[tid]
        nflag = len(FLAG_COLS)
        flags = list(C[5:5 + nflag])
        out.append([idlic, fonte, L[2], L[3], L[4], regime, pdoc, par[pdoc][1], vle, pct,
                    nota, T[4], T[3], T[2]] + flags + [C[2], cid, tid, C[5 + nflag]])
    return out


# --------------------------------------------------------------- writing
HDR_FILL = PatternFill('solid', fgColor='1F4E78')
HDR_FONT = Font(color='FFFFFF', bold=True, size=10)
KEY_FONT = Font(bold=True, color='1F4E78', size=10)
_THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
ONE_FILL = PatternFill('solid', fgColor='C6E4E0')
AUTO_FILL = PatternFill('solid', fgColor='FFE8CC')      # rows still to review
ZERO_FONT = Font(color='BFBFBF', size=10)


def _sheet(wb, title, headers, rows, widths, wrap=(), flag_span=None, review_col=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 32

    for r in rows:
        ws.append(list(r))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(vertical='top',
                                       wrap_text=cell.column_letter in wrap)
        row[0].font = KEY_FONT
        if flag_span:
            for c in range(flag_span[0], flag_span[1] + 1):
                cell = row[c - 1]
                if cell.value in (1, '1'):
                    cell.fill, cell.font = ONE_FILL, Font(bold=True, size=10)
                elif cell.value in (0, '0'):
                    cell.font = ZERO_FONT
                cell.alignment = Alignment(horizontal='center', vertical='center')
        if review_col and str(row[review_col - 1].value or '').startswith(ORIG_AUTO[:4]):
            row[review_col - 1].fill = AUTO_FILL

    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{ws.max_row}'
    ws.freeze_panes = 'A2'
    return ws


def _readme(wb, fonte, stats, extra):
    ws = wb.create_sheet('Leia-me', 0)
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 104
    nflag = len(FLAG_COLS)
    lines = [
        (f'MODELO DE CRITÉRIOS {fonte}', ''),
        ('', ''),
        ('Como está organizado', 'Esquema em estrela: um catálogo de regras + um catálogo de '
                                 'textos legais, ligados às licenças por uma tabela de factos. '
                                 'O ficheiro TUA e o ficheiro LURH têm EXACTAMENTE as mesmas '
                                 'folhas e as mesmas colunas.'),
        ('', ''),
        ('Catalogo_Criterios', f"{stats['cat']} regras. Uma linha por (Parâmetro, VLE, combinação "
                               f"de critérios). É AQUI que se corrige a interpretação de um "
                               f"critério — uma vez, não em centenas de linhas."),
        ('Textos_Legais', f"{stats['txt']} textos. Uma linha por (legislação, avaliação de "
                          'conformidade, interpretação). O parágrafo legal fica guardado uma só vez.'),
        ('Licencas_Criterios', f"{stats['fact']} linhas. A tabela de factos: para cada licença e "
                               'parâmetro, qual IdCriterio e qual IdTexto se aplicam.'),
        ('Dim_Licencas', f"{stats['lic']} licenças."),
        ('Dim_Parametros', f"{stats['par']} parâmetros. Nome no documento -> nome no catálogo "
                           '(vocabulário comum a TUA e LURH).'),
        ('Vista_Plana', f"{stats['fact']} linhas. Tudo junto numa tabela plana. "
                        'NÃO EDITAR: é regenerada a cada execução do script.'),
        ('', ''),
        ('IdLicenca', 'Nº TUA nos ficheiros TUA, Nº Licença nos ficheiros LURH.'),
        ('IdCriterio', f'{fonte}_PARAMETRO_n — ex.: {fonte}_CQO_1. O prefixo da fonte evita '
                       'colisões quando as duas tabelas são unidas.'),
        ('IdTexto', f'{fonte}_TX01 … {fonte}_TX{stats["txt"]:02d}.'),
        ('Fonte', "'TUA' ou 'LURH' em todas as tabelas — permite voltar a separar o modelo unido."),
        ('Regime', "'Normal' em todas as linhas TUA (o conceito não existe no TUA); "
                   "'Normal' / 'Ano de arranque' no LURH."),
        ('Origem dos critérios', 'Como foram atribuídos os critérios da regra. '
                                 f"'{ORIG_AUTO}' assinala classificação automática ainda por validar."),
        ('', ''),
        ('Como alterar um critério',
         '1) Editar a linha em Catalogo_Criterios. 2) Nada mais — todas as licenças que usam esse '
         'IdCriterio passam a refletir a alteração. 3) Voltar a correr o script apenas se quiser '
         'regenerar a Vista_Plana.'),
        ('Ligação ao Power BI',
         'Dim_Licencas[IdLicenca] 1—* Licencas_Criterios; Catalogo_Criterios[IdCriterio] 1—* '
         'Licencas_Criterios; Textos_Legais[IdTexto] 1—* Licencas_Criterios; '
         'Dim_Parametros[Parâmetro (documento)] 1—* Licencas_Criterios. Relações de sentido único.'),
        ('Unir TUA + LURH',
         'Append Queries sobre cada folha homónima dos dois ficheiros. As colunas coincidem uma a '
         'uma e os Ids têm prefixo de fonte, por isso a união não precisa de qualquer mapeamento. '
         'Segmentar por [Fonte] para ver só TUA ou só LURH.'),
    ] + list(extra) + [
        ('', ''),
        ('Gerado por', f'src/{fonte.lower()}/build_criteria_table.py + '
                       'extraction/criteria_model.py'),
        ('Gerado em', datetime.now().strftime('%Y-%m-%d %H:%M')),
    ]
    for a, b in lines:
        ws.append([a, b])
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=1).font = Font(bold=True, size=10)
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical='top')
    ws.sheet_view.showGridLines = False
    return ws


def write_workbook(rows, fonte, out_path, readme_extra=()):
    """Normalise `rows`, write the 7-sheet workbook, return the tables."""
    catalogo, textos, factos, licencas, parametros = normalise(rows, fonte)
    vista = flatten(catalogo, textos, factos, licencas, parametros)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    nflag = len(FLAG_COLS)

    _sheet(wb, 'Catalogo_Criterios', CATALOGO_COLS, catalogo,
           {'A': 20, 'B': 8, 'C': 18, 'D': 20, 'E': 11,
            **{get_column_letter(i): 12 for i in range(6, 6 + nflag)},
            get_column_letter(6 + nflag): 26, get_column_letter(7 + nflag): 12},
           flag_span=(6, 5 + nflag), review_col=6 + nflag)

    _sheet(wb, 'Textos_Legais', TEXTOS_COLS, textos,
           {'A': 14, 'B': 8, 'C': 46, 'D': 92, 'E': 60, 'F': 12},
           wrap=('C', 'D', 'E'))

    _sheet(wb, 'Licencas_Criterios', FACTOS_COLS, factos,
           {'A': 22, 'B': 8, 'C': 16, 'D': 42, 'E': 11, 'F': 14, 'G': 12, 'H': 20, 'I': 14})

    _sheet(wb, 'Dim_Licencas', LICENCAS_COLS, licencas,
           {'A': 22, 'B': 8, 'C': 32, 'D': 12, 'E': 62})

    _sheet(wb, 'Dim_Parametros', PARAMETROS_COLS, parametros,
           {'A': 48, 'B': 22, 'C': 8})

    vflag = VISTA_COLS.index(FLAG_COLS[0]) + 1
    _sheet(wb, 'Vista_Plana', VISTA_COLS, vista,
           {'A': 22, 'B': 8, 'C': 28, 'D': 10, 'E': 46, 'F': 14, 'G': 34, 'H': 16, 'I': 10,
            'J': 14, 'K': 11, 'L': 46, 'M': 60, 'N': 40,
            **{get_column_letter(i): 11 for i in range(vflag, vflag + nflag)},
            get_column_letter(vflag + nflag): 16,
            get_column_letter(vflag + nflag + 1): 20,
            get_column_letter(vflag + nflag + 2): 14,
            get_column_letter(vflag + nflag + 3): 26},
           wrap=('L', 'M', 'N'), flag_span=(vflag, vflag + nflag - 1))

    _readme(wb, fonte, {'cat': len(catalogo), 'txt': len(textos), 'fact': len(factos),
                        'lic': len(licencas), 'par': len(parametros)}, readme_extra)
    wb.save(out_path)
    return catalogo, textos, factos, licencas, parametros, vista


def check_integrity(catalogo, textos, factos, licencas, parametros):
    """Referential integrity of the produced model -> list of problems."""
    problems = []
    cids = {r[0] for r in catalogo}
    tids = {r[0] for r in textos}
    lids = {r[0] for r in licencas}
    pids = {r[0] for r in parametros}
    if len(cids) != len(catalogo):
        problems.append('IdCriterio is not unique')
    if len(tids) != len(textos):
        problems.append('IdTexto is not unique')
    for label, missing in (('IdCriterio', {f[7] for f in factos} - cids),
                           ('IdTexto', {f[8] for f in factos} - tids),
                           ('IdLicenca', {f[0] for f in factos} - lids),
                           ('Parâmetro', {f[3] for f in factos} - pids)):
        if missing:
            problems.append(f'{len(missing)} orphan {label} reference(s): {sorted(missing)[:4]}')
    for label, unused in (('catalogue', cids - {f[7] for f in factos}),
                          ('text', tids - {f[8] for f in factos})):
        if unused:
            problems.append(f'{len(unused)} unused {label} row(s)')
    nflag = len(FLAG_COLS)
    if sum(r[6 + nflag] for r in catalogo) != len(factos):
        problems.append('catalogue usage counts do not sum to the fact row count')
    if sum(r[5] for r in textos) != len(factos):
        problems.append('text usage counts do not sum to the fact row count')
    return problems


def backup(path):
    """Timestamped .bak_ copy, same convention as sync_criterios.py."""
    import shutil
    if os.path.exists(path):
        bak = f"{path}.bak_{datetime.now():%Y%m%d_%H%M%S}"
        shutil.copy2(path, bak)
        return bak
    return None
