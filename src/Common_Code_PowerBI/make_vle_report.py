# -*- coding: utf-8 -*-
"""One sheet with, for every plant and parameter: the limit, how often it must
be analysed, and until when the licence that imposes it is valid.

Joins the three tables the licences produce:
    Licenses      identity of the plant and validity of its licence
    Conditions    the limit (VLE) per parameter
    Autocontrolo  sampling frequency and number of analyses per year

Sheets
    VLE_por_ETAR  one row per (plant, parameter, period)
    Resumo        one row per plant: parameters, analyses/year, validity

Run:  python src/make_vle_report.py
Output: outputs/reports/VLE_por_ETAR.xlsx
"""
import datetime
import os
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))   # project root (this module is in src/Common_Code_PowerBI)
FEED = os.getenv('EPAL_POWERBI_OUT', os.path.join(ROOT, 'data', 'powerbi'))
OUT_DIR = os.path.join(ROOT, 'outputs', 'reports')
OUT = os.path.join(OUT_DIR, 'VLE_por_ETAR.xlsx')

WINDOW_MONTHS = 8


def read(name):
    import openpyxl
    path = os.path.join(FEED, f'{name}.xlsx')
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    it = wb.active.values
    try:
        hdr = [str(c).strip() if c is not None else '' for c in next(it)]
    except StopIteration:
        wb.close()
        return []
    out = [dict(zip(hdr, r)) for r in it]
    wb.close()
    return out


def parse_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        for f in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
            try:
                return datetime.datetime.strptime(v[:10], f).date()
            except ValueError:
                pass
    return None


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    m += 1
    leap = y % 4 == 0 and (y % 100 or y % 400 == 0)
    day = min(d.day, [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return datetime.date(y, m, day)


def txt(v):
    return '' if v is None else str(v).strip()


def main():
    today = datetime.date.today()
    limit = add_months(today, WINDOW_MONTHS)

    licences = {}
    for r in read('Licenses'):
        tua = txt(r.get('Nº TUA'))
        if tua:
            licences[tua] = r
    if not licences:
        print('! Licenses.xlsx vazio — correr make_powerbi_all.py primeiro', file=sys.stderr)
        return 1

    # sampling obligations per (licence, parameter, period); rows are duplicated
    # in the source, so the same obligation is only counted once
    auto = {}
    for a in read('Autocontrolo'):
        key = (txt(a.get('Nº TUA')), txt(a.get('Parametro')), txt(a.get('Periodo')) or 'Normal')
        rec = auto.setdefault(key, {'freq': '', 'tipo': '', 'n': 0, 'local': '', 'seen': set()})
        sig = (txt(a.get('Frequência de amostragem')), txt(a.get('Nº análises requeridas')))
        if sig in rec['seen']:
            continue
        rec['seen'].add(sig)
        rec['freq'] = rec['freq'] or txt(a.get('Frequência de amostragem'))
        rec['tipo'] = rec['tipo'] or txt(a.get('Tipo de amostragem'))
        rec['local'] = rec['local'] or txt(a.get('Local de amostragem'))
        try:
            rec['n'] += int(float(txt(a.get('Nº análises requeridas')) or 0))
        except ValueError:
            pass

    rows = []
    for c in read('Conditions'):
        tua = txt(c.get('Nº TUA'))
        lic = licences.get(tua)
        if not lic:
            continue
        param = txt(c.get('Parametro'))
        period = txt(c.get('Periodo')) or 'Normal'
        obl = auto.get((tua, param, period)) or auto.get((tua, param, 'Normal')) or {}

        validade = parse_date(lic.get('Validade (ISO)')) or parse_date(lic.get('Data de Validade'))
        dias = (validade - today).days if validade else None
        if validade is None:
            estado = 'Sem data'
        elif validade < today:
            estado = 'Caducada'
        elif validade <= limit:
            estado = f'A expirar (< {WINDOW_MONTHS} meses)'
        else:
            estado = 'Válida'

        rows.append({
            'Nº TUA': tua,
            'Regime': txt(lic.get('TUA/LURH')),
            'ETAR': txt(lic.get('Estabelecimento')),
            'Designação': txt(lic.get('Designação')),
            'Código APA': txt(lic.get('Código APA')),
            'Código TURH': txt(lic.get('Código TURH')),
            'Município': txt(lic.get('Município')),
            'Região': txt(lic.get('REGIÃO')),
            'Validade da licença': validade.strftime('%d/%m/%Y') if validade else '',
            'Dias p/ expirar': dias if dias is not None else '',
            'Estado da licença': estado,
            'Parâmetro': param,
            'Unidade': txt(c.get('Unidade')),
            'Período': period,
            'VLE': txt(c.get('VLE')),
            'VLE máx': txt(c.get('VLE máx')),
            '% mín. redução': txt(c.get('VLE (% mín. redução)')),
            'Frequência de amostragem': obl.get('freq', ''),
            'Tipo de amostragem': obl.get('tipo', ''),
            'Local de amostragem': obl.get('local', ''),
            'Nº análises/ano': obl.get('n', 0) or '',
            'Parâmetro (texto da licença)': txt(c.get('Parâmetro')),
        })

    rows.sort(key=lambda r: (r['ETAR'], r['Parâmetro'], r['Período']))

    # per-plant summary
    per = defaultdict(lambda: {'params': set(), 'analises': 0})
    for r in rows:
        p = per[r['Nº TUA']]
        p['params'].add(r['Parâmetro'])
        p['analises'] += r['Nº análises/ano'] if isinstance(r['Nº análises/ano'], int) else 0
        p['row'] = r

    resumo = []
    for tua, v in per.items():
        r = v['row']
        resumo.append({
            'Nº TUA': tua, 'Regime': r['Regime'], 'ETAR': r['ETAR'],
            'Código APA': r['Código APA'], 'Código TURH': r['Código TURH'],
            'Município': r['Município'], 'Região': r['Região'],
            'Validade da licença': r['Validade da licença'],
            'Dias p/ expirar': r['Dias p/ expirar'],
            'Estado da licença': r['Estado da licença'],
            'Nº parâmetros': len(v['params']),
            'Nº análises/ano': v['analises'],
        })
    resumo.sort(key=lambda r: (r['Estado da licença'], r['ETAR']))

    write(rows, resumo, today)
    print(f'  {len(rows)} linhas (ETAR x parâmetro), {len(resumo)} ETAR')
    print(f'  -> {os.path.relpath(OUT, ROOT)}')
    return 0


def write(rows, resumo, today):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()

    head_fill = PatternFill('solid', fgColor='0D6ABF')
    head_font = Font(color='FFFFFF', bold=True)
    colours = {'Caducada': 'F5B7B1', 'Válida': 'A9DFBF', 'Sem data': 'D5D8DC'}

    def sheet(ws, data, name, widths):
        cols = list(data[0].keys())
        ws.append(cols)
        for r in data:
            ws.append([r.get(c, '') for c in cols])
        for cell in ws[1]:
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        ws.freeze_panes = 'A2'
        last = max(ws.max_row, 2)
        t = Table(displayName=name, ref=f'A1:{get_column_letter(len(cols))}{last}')
        t.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True)
        ws.add_table(t)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        # colour the licence state so an expired licence is visible at a glance
        if 'Estado da licença' in cols:
            idx = cols.index('Estado da licença') + 1
            for row in range(2, last + 1):
                v = str(ws.cell(row=row, column=idx).value or '')
                fill = colours.get(v) or ('FAD7A0' if v.startswith('A expirar') else None)
                if fill:
                    ws.cell(row=row, column=idx).fill = PatternFill('solid', fgColor=fill)

    ws1 = wb.active
    ws1.title = 'VLE_por_ETAR'
    sheet(ws1, rows or [{'Nº TUA': ''}], 'VLE_por_ETAR',
          (20, 8, 30, 24, 14, 20, 18, 14, 14, 10, 20, 14, 14, 10, 14, 12, 12, 16, 18, 22, 12, 34))

    ws2 = wb.create_sheet('Resumo')
    sheet(ws2, resumo or [{'Nº TUA': ''}], 'Resumo_ETAR',
          (20, 8, 30, 14, 20, 18, 14, 14, 10, 20, 12, 14))

    ws3 = wb.create_sheet('Leia-me')
    for line in (
            'VLE por ETAR — gerado automaticamente',
            f'Data de geração: {today:%d/%m/%Y}',
            '',
            'VLE_por_ETAR   uma linha por ETAR, parâmetro e período.',
            'Resumo         uma linha por ETAR: nº de parâmetros e de análises por ano.',
            '',
            'Origem dos dados',
            '  Identificação e validade  ->  Licenses (extraído das licenças)',
            '  VLE                       ->  Conditions',
            '  Frequência e nº análises  ->  Autocontrolo',
            '',
            'Período: "Estiagem" indica o valor-limite mais exigente aplicável ao',
            'período de estiagem; "Normal" o valor aplicável no resto do ano.',
            '',
            'Este ficheiro é gerado por src/make_vle_report.py. Não editar à mão:',
            'a próxima execução substitui-o.'):
        ws3.append([line])
    ws3.column_dimensions['A'].width = 78
    ws3['A1'].font = Font(bold=True, size=13)

    wb.save(OUT)


if __name__ == '__main__':
    sys.exit(main())
