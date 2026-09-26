# -*- coding: utf-8 -*-
"""Recreate the AdVT_DepDOS_VLE.xlsx layout (Anexo III/IV/V/VI-b PCQAR AdVT),
filled from the project's own extracted licences instead of hand-maintained.

Four sheets, one per AdVT department (matches TabelaCentroide_ETAR's own
'Departamento DOS' values), each a wide table: one row per licensed ETAR,
one column per discharge parameter (VLE), pivoted from data/powerbi/Conditions.xlsx.

    DCA_S  Centro Alentejo   (Anexo III-b)
    DNA_S  Norte Alentejo    (Anexo IV-b)
    DBB_S  Beira Baixa       (Anexo V-b)
    DBA_S  Beira Alta        (Anexo VI-b)

Column set per sheet = the union of the original template's parameter columns
for that department AND any parameter actually present among that
department's licences (so a newly-extracted licence with a parameter never
seen before still gets a column, instead of being silently dropped).

Adds one column not in the original template: "Período de estiagem" — the
dry-season date range, when the licence's own conditions define one (e.g.
"considera-se o período de estiagem de 1 de junho a 30 de setembro"). Left
blank when no such clause exists. This is independent of the per-parameter
VLE columns, which already show the estiagem-specific limit inline (second
line) when the licence sets one, exactly as the original file did.

Run:  python src/Common_Code_PowerBI/make_advt_depdos_vle.py
Output: outputs/reports/AdVT_DepDOS_VLE.xlsx
"""
import datetime
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))   # project root (this module is in src/Common_Code_PowerBI)
FEED = os.getenv('EPAL_POWERBI_OUT', os.path.join(ROOT, 'data', 'powerbi'))
REGISTRY_XLSX = os.path.join(FEED, 'TabelaCentroide_ETAR_newDep_sensity.xlsx')
TUA_MASTER = os.path.join(ROOT, 'src', 'Extraction_Code_TUA', 'master_tua.xlsx')
LURH_MASTER = os.path.join(ROOT, 'src', 'Extraction_Code_LURH', 'master_lurh.xlsx')
OUT_DIR = os.path.join(ROOT, 'outputs', 'reports')
OUT = os.path.join(OUT_DIR, 'AdVT_DepDOS_VLE.xlsx')

# Department -> (sheet name, Anexo label, title suffix) — matches the original file exactly.
DEPARTAMENTOS = {
    'Centro Alentejo': ('DCA_S', 'III-b', 'aplicáveis a DCA-S'),
    'Norte Alentejo':  ('DNA_S', 'IV-b',  'aplicáveis a DNA-S'),
    'Beira Baixa':     ('DBB_S', 'V-b',   'aplicáveis ao subsistema de saneamento de AdVT/DBB-S'),
    'Beira Alta':      ('DBA_S', 'VI-b',  'aplicáveis ao subsistema de saneamento de AdVT/BA'),
}

# Parameter key (as it appears in Conditions.xlsx's 'Parametro' column) -> the
# original template's column header text for that parameter (kept verbatim,
# including the line breaks, so the recreated file matches the source layout).
PARAM_HEADERS = {
    'pH':       'pH\n(Escala Sorensen)',
    'CBO5':     'CBO5 \n(mg/L O2)',
    'CQO':      'CQO\n(mg/L O2)',
    'SST':      'SST\n(mg/L)',
    'O&G':      'OG\n(mg/L)',
    'Nt':       'Nt\n(mg/L N)',
    'Pt':       'Pt\n(mg/L P)',
    'N-NH4':    'NH4 (mg/L NH4)',
    'E. coli':  'EC\n(NMP/ 100mL)',
    'CF':       'CF\n(NMP/ 100mL)',
    'Ovos de parasitas intestinais': 'Ovos PI (N/L)',
    'Sulfuretos': 'S\n(mg/L S)',
    'Cr total': 'Cr\n(mg/L Cr)',
    'Fenóis':   'Fenóis (mg/L C6H5OH)',
    # Present in the original template but not seen in any extracted licence
    # yet — kept as columns (matching the layout) so a future licence that
    # sets one of these has somewhere to land; the Parametro key each maps
    # from is a guess based on the extraction's own abbreviation scheme.
    'Detergentes':      'Det (mg/L)',
    'Sulfatos':         'Sulfatos (mg/L SO4)',
    'Óleos minerais':   'OM\n(mg/L)',
}
# 'Caudal' is a flow measurement, not a discharge VLE, so it has no column
# here (the original template doesn't carry it either).
EXCLUDED_PARAMS = {'Caudal'}

# The original template's per-department column set and order (left to right,
# after the fixed 'Título de licenciamento' columns). Anything a department's
# licences use beyond this set is appended automatically (see build_columns()).
TEMPLATE_COLUMNS = {
    'Centro Alentejo': ['pH', 'CBO5', 'CQO', 'SST', 'O&G', 'Nt', 'Pt', 'N-NH4', 'E. coli', 'CF',
                         'Ovos de parasitas intestinais', 'Sulfuretos', 'Cr total'],
    'Norte Alentejo':  ['pH', 'CBO5', 'CQO', 'SST', 'O&G', 'Nt', 'Pt', 'N-NH4', 'E. coli', 'CF',
                         'Sulfuretos', 'Cr total', 'Fenóis', 'Detergentes', 'Sulfatos'],
    'Beira Baixa':     ['pH', 'CBO5', 'CQO', 'SST', 'O&G', 'Nt', 'Pt', 'N-NH4', 'E. coli'],
    'Beira Alta':      ['pH', 'CBO5', 'CQO', 'SST', 'O&G', 'Nt', 'Pt', 'N-NH4', 'CF', 'E. coli',
                         'Óleos minerais'],
}

DATE_RANGE_RE = re.compile(
    r'per[ií]odo\s+de\s+estiagem\D{0,60}?'
    r'(\d{1,2}\s*(?:de\s+)?[a-zçãéóú]+)\s*(?:a|e)\s*(\d{1,2}\s*(?:de\s+)?[a-zçãéóú]+)',
    re.IGNORECASE,
)


def txt(v):
    return '' if v is None else str(v).strip()


def num(v):
    if v is None or v == '':
        return None
    try:
        f = float(str(v).replace(',', '.'))
        return f
    except (ValueError, TypeError):
        return None


def fmt_vle(v):
    """VLE cells from Conditions.xlsx are already the licence's own text
    ('6,0-9,0', '150', ...) — pass through, just normalise blank/NaN."""
    s = txt(v)
    return s


def read_sheet_rows(path, sheet_name=None):
    import openpyxl
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    it = ws.values
    try:
        hdr = [str(c).strip() if c is not None else '' for c in next(it)]
    except StopIteration:
        wb.close()
        return []
    out = [dict(zip(hdr, r)) for r in it]
    wb.close()
    return out


def extract_estiagem_period(text):
    m = DATE_RANGE_RE.search(text or '')
    if not m:
        return None
    d1 = re.sub(r'^0+(?=\d)', '', m.group(1).strip())
    d2 = re.sub(r'^0+(?=\d)', '', m.group(2).strip())
    return f'{d1} a {d2}'


def find_estiagem_periods():
    """Nº de licença -> dry-season date range text, from whichever 'Outras
    Condições' clause defines one. Scans both regimes' own condition-clause
    sheets (TUA's 'Outras Condicoes (3.21)', LURH's 'CondicoesTexto')."""
    periods = {}
    for r in read_sheet_rows(TUA_MASTER, 'Outras Condicoes (3.21)'):
        p = extract_estiagem_period(txt(r.get('Condição')))
        if p:
            periods.setdefault(txt(r.get('Nº TUA')), p)
    for r in read_sheet_rows(LURH_MASTER, 'CondicoesTexto'):
        if txt(r.get('Secção')) != 'Outras':
            continue
        p = extract_estiagem_period(txt(r.get('Condição')))
        if p:
            periods.setdefault(txt(r.get('Nº Licença')), p)
    return periods


def read_conditions():
    """Conditions.xlsx pivoted: Nº TUA -> Parametro -> {'Normal': vle, 'Estiagem': vle}.

    Shared with make_licencas_parametros_regiao.py, so both reports see the
    same limits — including the correction below."""
    cond = {}
    for c in read_sheet_rows(os.path.join(FEED, 'Conditions.xlsx')):
        tua = txt(c.get('Nº TUA'))
        param = txt(c.get('Parametro'))
        if not tua or not param or param in EXCLUDED_PARAMS:
            continue
        periodo = txt(c.get('Periodo')) or 'Normal'
        cond.setdefault(tua, {}).setdefault(param, {})[periodo] = fmt_vle(c.get('VLE'))

    # Known upstream extraction defect (LURH L013710.2021.RH5A, ETAR Entroncamento):
    # the source PDF's ruled table glued "Sólidos Suspensos Totais (período húmido)
    # (mg/L) 35" and "Óleos e Gorduras (mg/L)" into one cell with a single VLE of
    # 15 — extract_lurh.py reads that as SST/Normal=15 and drops O&G entirely.
    # Corrected here from the original AdVT_DepDOS_VLE.xlsx (SST=35, O&G=15),
    # pending a proper fix in extract_lurh.py's ruled-table cell splitting.
    _bad = cond.get('L013710.2021.RH5A', {})
    if _bad.get('SST', {}).get('Normal') == '15':
        _bad['SST']['Normal'] = '35'
        _bad.setdefault('O&G', {})['Normal'] = '15'
    return cond


def licence_columns(tua, lic, reg, estiagem_periods):
    """The eight identity/validity columns both reports share, for one licence."""
    return {
        'Área do CO': re.sub(r'^CO\s+(?:de|do|da)?\s*', '', txt(reg.get('Centro_Operacional'))).replace('/', '-'),
        'Município': txt(lic.get('Município')) or txt(reg.get('CONCELHO')),
        'Infraestrutura': txt(lic.get('Estabelecimento')),
        'N.º TUA': tua if str(tua).startswith('TUA') else '',
        'N.º LURH': txt(lic.get('Código TURH')) if str(tua).startswith('TUA') else tua,
        'Data de início': lic.get('Data de Entrada em Vigor') or lic.get('Data de Emissão'),
        'Data de validade': lic.get('Data de Validade'),
        'Período de estiagem': estiagem_periods.get(tua, ''),
    }


def vle_cell(by_period):
    """One parameter's limit as a single cell: the two values stacked when the
    licence sets a different one for the dry season, otherwise just the one."""
    normal = by_period.get('Normal', '')
    estiagem = by_period.get('Estiagem', '')
    if normal and estiagem and normal != estiagem:
        return f'{normal}\n(período húmido*)\n{estiagem}\n(período estiagem*)'
    return normal or estiagem


def build_columns(dept, params_seen):
    """Template order first, then any parameter this department actually
    uses that the template didn't already list (union, not intersection —
    nothing a real licence sets is ever dropped)."""
    cols = list(TEMPLATE_COLUMNS[dept])
    for p in sorted(params_seen - set(cols) - EXCLUDED_PARAMS):
        if p in PARAM_HEADERS:
            cols.append(p)
        else:
            PARAM_HEADERS[p] = p  # unknown parameter: fall back to its own key as the header
            cols.append(p)
    return cols


def main():
    licences = {r.get('Nº TUA'): r for r in read_sheet_rows(os.path.join(FEED, 'Licenses.xlsx')) if r.get('Nº TUA')}
    if not licences:
        print('! Licenses.xlsx vazio — correr make_powerbi_all.py primeiro', file=sys.stderr)
        return 1

    registry = {r.get('CODMAXIMO'): r for r in read_sheet_rows(REGISTRY_XLSX) if r.get('CODMAXIMO')}

    estiagem_periods = find_estiagem_periods()

    cond = read_conditions()
    params_by_dept = {}

    rows_by_dept = {d: [] for d in DEPARTAMENTOS}
    unclassified = []

    for tua, lic in licences.items():
        cod = lic.get('CODMAXIMO')
        reg = registry.get(cod) if cod else None
        dept = txt(reg.get('Departamento DOS')) if reg else ''
        if dept not in DEPARTAMENTOS:
            unclassified.append(lic)
            continue

        params_by_dept.setdefault(dept, set()).update(cond.get(tua, {}).keys())

        row = licence_columns(tua, lic, reg, estiagem_periods)
        row.update({'_dept': dept, '_tua': tua})
        for param, by_period in cond.get(tua, {}).items():
            row[param] = vle_cell(by_period)
        rows_by_dept[dept].append(row)

    for d in rows_by_dept:
        rows_by_dept[d].sort(key=lambda r: (r['Área do CO'], r['Município'], r['Infraestrutura']))

    write(rows_by_dept, params_by_dept, unclassified)
    total = sum(len(v) for v in rows_by_dept.values())
    print(f'{total} licenças classificadas em {len(DEPARTAMENTOS)} departamentos; '
          f'{len(unclassified)} sem Departamento DOS (ver aba "Não classificado")')
    for d, (sheet, *_ ) in DEPARTAMENTOS.items():
        print(f'  {sheet} ({d}): {len(rows_by_dept[d])} licenças, '
              f'{len(build_columns(d, params_by_dept.get(d, set())))} colunas de parâmetro')
    print(f'  -> {os.path.relpath(OUT, ROOT)}')
    return 0


def write(rows_by_dept, params_by_dept, unclassified):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HEAD_FILL_LIC = PatternFill('solid', fgColor='1F4E78')
    HEAD_FILL_VLE = PatternFill('solid', fgColor='2E7D32')
    HEAD_FILL_NEW = PatternFill('solid', fgColor='B7791F')
    HEAD_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=9)
    BODY_FONT = Font(name='Arial', size=9)
    THIN = Side(style='thin', color='D9D9D9')
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    FIXED_COLS = ['Área do CO', 'Município', 'Infraestrutura', 'N.º TUA', 'N.º LURH',
                  'Data de início', 'Data de validade', 'Período de estiagem']

    for dept, (sheet_name, anexo, suffix) in DEPARTAMENTOS.items():
        ws = wb.create_sheet(sheet_name)
        param_cols = build_columns(dept, params_by_dept.get(dept, set()))
        cols = FIXED_COLS + param_cols

        ws.cell(1, len(cols), f'Anexo {anexo} PCQAR AdVT — Requisitos de descarga de águas residuais {suffix}').font = Font(name='Arial', bold=True, size=10)
        ws.cell(2, len(cols), f'Gerado automaticamente a partir das licenças extraídas — {datetime.date.today():%d/%m/%Y}').font = Font(name='Arial', italic=True, size=8)

        r0 = 4
        # Matches the original layout: "Título de licenciamento" spans only the
        # licence-number/date columns (Área do CO / Município / Infraestrutura,
        # and the new Período de estiagem column, sit under it unlabeled — same
        # as in the source file).
        lic_start = FIXED_COLS.index('N.º TUA') + 1
        lic_end = FIXED_COLS.index('Data de validade') + 1
        ws.cell(r0, lic_start, 'Título de licenciamento').font = HEAD_FONT
        ws.cell(r0, lic_start).fill = HEAD_FILL_LIC
        ws.merge_cells(start_row=r0, start_column=lic_start, end_row=r0, end_column=lic_end)
        vle_start = len(FIXED_COLS) + 1
        vle_end = vle_start + max(len(param_cols) - 1, 0)
        ws.cell(r0, vle_start, 'Valores Limite de Emissão (VLE)').font = HEAD_FONT
        ws.cell(r0, vle_start).fill = HEAD_FILL_VLE
        if vle_end > vle_start:
            ws.merge_cells(start_row=r0, start_column=vle_start, end_row=r0, end_column=vle_end)

        header_row = r0 + 1
        for j, col in enumerate(cols, start=1):
            c = ws.cell(header_row, j, PARAM_HEADERS.get(col, col) if j > len(FIXED_COLS) else col)
            c.font = HEAD_FONT
            c.fill = HEAD_FILL_NEW if col == 'Período de estiagem' else (HEAD_FILL_VLE if j > len(FIXED_COLS) else HEAD_FILL_LIC)
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = BORDER
        ws.row_dimensions[header_row].height = 32

        data = rows_by_dept[dept]
        for i, row in enumerate(data, start=header_row + 1):
            for j, col in enumerate(cols, start=1):
                v = row.get(col, '')
                if col in ('Data de início', 'Data de validade') and isinstance(v, datetime.datetime):
                    v = v.date()
                c = ws.cell(i, j, v if v not in ('', None) else None)
                c.font = BODY_FONT
                c.border = BORDER
                c.alignment = Alignment(wrap_text=True, vertical='center')
                if col in ('Data de início', 'Data de validade') and v:
                    c.number_format = 'DD/MM/YYYY'
            if (i - header_row) % 2 == 0:
                for j in range(1, len(cols) + 1):
                    ws.cell(i, j).fill = PatternFill('solid', fgColor='F2F6FA')

        widths = {'Área do CO': 22, 'Município': 16, 'Infraestrutura': 24, 'N.º TUA': 18,
                  'N.º LURH': 20, 'Data de início': 12, 'Data de validade': 12,
                  'Período de estiagem': 20}
        for j, col in enumerate(cols, start=1):
            ws.column_dimensions[get_column_letter(j)].width = widths.get(col, 12)
        ws.freeze_panes = ws.cell(header_row + 1, len(FIXED_COLS) + 1).coordinate

    if unclassified:
        ws = wb.create_sheet('Não classificado')
        ws.append(['Nº TUA', 'Estabelecimento', 'CODMAXIMO', 'Motivo'])
        for c in ws[1]:
            c.font = HEAD_FONT
            c.fill = HEAD_FILL_NEW
        for lic in unclassified:
            ws.append([lic.get('Nº TUA'), lic.get('Estabelecimento'), lic.get('CODMAXIMO'),
                       'Sem Departamento DOS na tabela de registo (TabelaCentroide_ETAR_newDep_sensity.xlsx)'])
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 24
        ws.column_dimensions['D'].width = 60

    ws = wb.create_sheet('Leia-me', 0)
    for line in (
            'AdVT — Requisitos de descarga (VLE) por departamento — gerado automaticamente',
            f'Data de geração: {datetime.date.today():%d/%m/%Y}',
            '',
            'Recria o layout do ficheiro AdVT_DepDOS_VLE.xlsx original, uma aba por',
            'departamento (DCA_S / DNA_S / DBB_S / DBA_S), preenchida a partir das',
            'licenças TUA e LURH já extraídas — deixa de ser preenchido à mão.',
            '',
            'Departamento de cada ETAR: TabelaCentroide_ETAR_newDep_sensity.xlsx,',
            'coluna "Departamento DOS", associada à licença pelo CODMAXIMO já',
            'atribuído em Licenses.xlsx (ver src/Common_Code_PowerBI/codmaximo_match.py).',
            '',
            'Colunas de parâmetro: união das colunas do modelo original com quaisquer',
            'parâmetros que apareçam nas licenças desse departamento e ainda não',
            'tivessem coluna — nenhum parâmetro extraído fica de fora.',
            '',
            'Quando uma licença define um VLE diferente para o período de estiagem,',
            'a célula do parâmetro mostra os dois valores (húmido / estiagem), tal',
            'como no ficheiro original.',
            '',
            'Coluna nova "Período de estiagem": as datas do período seco, quando a',
            'própria licença as define numa cláusula (ex. "considera-se o período de',
            'estiagem de 1 de junho a 30 de setembro"). Em branco quando a licença',
            'não define um período — não confundir com o valor de VLE de estiagem,',
            'que continua na coluna do parâmetro.',
            '',
            'Licenças sem Departamento DOS na tabela de registo aparecem na aba',
            '"Não classificado" em vez de serem descartadas.',
            '',
            'Correção manual conhecida: ETAR Entroncamento (L013710.2021.RH5A, aba',
            'DBB_S) — a extração LURH junta "SST" e "OG" numa só célula na origem;',
            'SST/OG foram corrigidos aqui (35 / 15, confirmados no ficheiro original)',
            'até isso ser corrigido em extract_lurh.py.',
            '',
            'Este ficheiro é gerado por src/Common_Code_PowerBI/make_advt_depdos_vle.py.',
            'Não editar à mão: a próxima execução substitui-o.'):
        ws.append([line])
    ws.column_dimensions['A'].width = 90
    ws['A1'].font = Font(name='Arial', bold=True, size=13)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name='Arial', size=10)

    wb.save(OUT)


if __name__ == '__main__':
    sys.exit(main())
