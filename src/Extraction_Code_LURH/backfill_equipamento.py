# -*- coding: utf-8 -*-
"""backfill_equipamento.py — preenche a folha 'Equipamento' do master a partir
das licenças já extraídas, sem reprocessar os PDFs.

A partir daqui o extract_lurh.py mantém a folha sozinho (build_sheets emite
'Equipamento' em cada extração); este script existe só para a carga inicial das
licenças que já foram extraídas antes da folha existir.

    python backfill_equipamento.py                       # data/extracted -> master_lurh.xlsx
    python backfill_equipamento.py <pasta> [master.xlsx] [--all]

Só escreve licenças que já existam na folha Licenses, para o modelo do Power BI
não ganhar linhas órfãs; as restantes são listadas no fim (são extrações que
nunca chegaram ao master, tipicamente runs em que ele estava aberto no Excel).
Use --all para as incluir à mesma.

Lê o workbook '*_extracted.xlsx' de cada licença — ETAR em Resumo!A1, Nº Licença
no par etiqueta/valor do Resumo, e 'Equipamento de controlo' no bloco de rodapé
da folha Autocontrolo. Cai para o JSON correspondente quando o workbook não
tiver o campo.
"""
import json
import os
import re
import sys

import openpyxl

SHEET = 'Equipamento'
HEADERS = ['ETAR', 'Nº Licença', 'Equipamento de controlo']

LIC_RE = re.compile(r'[LP]\d{6}[._]\d{4}[._]RH[0-9A-Za-z]+')
EQUIP_RE = re.compile(r'equipamento\s+de\s+controlo', re.I)
NUM_RE = re.compile(r'^\s*n[ºo°]?\.?\s*licen[çc]a\s*$', re.I)


def _pair(ws, pattern):
    """Valor à direita da primeira etiqueta que corresponde ao padrão.

    Os workbooks por licença dispõem os campos de cabeçalho e de rodapé como
    pares etiqueta | valor, por isso o valor é a célula seguinte não vazia.
    """
    if ws is None:
        return None
    for row in ws.iter_rows(max_col=2):
        for cell in row:
            if isinstance(cell.value, str) and pattern.search(cell.value):
                for c in range(cell.column + 1, cell.column + 4):
                    v = ws.cell(cell.row, c).value
                    if v is not None and str(v).strip():
                        return str(v).strip()
                return ''
    return None


def from_workbook(path):
    """(etar, numero, equipamento) — equipamento é None se a etiqueta faltar."""
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        resumo = next((w for w in wb.worksheets
                       if w.title.strip().lower().startswith('resumo')), None)
        auto = next((w for w in wb.worksheets
                     if 'autocontrolo' in w.title.strip().lower()), None)
        etar = (resumo['A1'].value if resumo is not None else None) or ''
        numero = _pair(resumo, NUM_RE)
        if not numero:
            for w in wb.worksheets:
                numero = _pair(w, NUM_RE)
                if numero:
                    break
        return str(etar).strip(), numero, _pair(auto, EQUIP_RE)
    finally:
        wb.close()


def from_json(path):
    """Mesmos três campos a partir do JSON — a fonte de que o workbook é escrito."""
    with open(path, encoding='utf-8') as fh:
        d = json.load(fh)
    g = d.get('dados_gerais', {}) or {}
    return ((d.get('tratamento', {}) or {}).get('Designação', ''),
            g.get('Nº Licença') or d.get('file_name', ''),
            ((d.get('autocontrolo', {}) or {}).get('Equipamento de controlo') or '').strip())


def collect(folder):
    records, sem_campo = [], []
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith('.xlsx') or name.startswith('~$'):
            continue
        path = os.path.join(folder, name)
        try:
            etar, numero, equip = from_workbook(path)
        except Exception as exc:
            print(f'  ! {name}: não foi possível ler ({exc})')
            continue

        # O workbook não trouxe o campo — tenta o JSON irmão antes de desistir.
        jpath = os.path.splitext(path)[0] + '.json'
        if (equip is None or not numero) and os.path.isfile(jpath):
            j_etar, j_num, j_equip = from_json(jpath)
            etar, numero = etar or j_etar, numero or j_num
            if equip is None:
                equip = j_equip

        if not numero:
            m = LIC_RE.search(name)
            numero = m.group(0).replace('_', '.') if m else None
        if not numero:
            print(f'  ! {name}: nº de licença não identificado')
            continue
        if equip is None:
            equip = ''
        if not equip:
            sem_campo.append(numero)
        records.append((etar, numero, equip))
    return records, sem_campo


def licencas_do_master(master_path):
    """Chaves da folha Licenses — o conjunto de licenças que o master reconhece."""
    wb = openpyxl.load_workbook(master_path, read_only=True)
    try:
        ws = wb['Licenses']
        header = [c.value for c in next(ws.iter_rows(max_row=1))]
        col = header.index('Nº Licença') + 1
        return {str(r[col - 1]).strip() for r in ws.iter_rows(min_row=2, values_only=True)
                if r[col - 1]}
    finally:
        wb.close()


def upsert(master_path, records):
    """Escreve na folha Equipamento, emparelhando por Nº Licença.

    Usa openpyxl e não pandas, para que as restantes folhas do master mantenham
    conteúdo e formatação. Devolve (atualizadas, adicionadas, caminho).
    """
    wb = openpyxl.load_workbook(master_path)
    if SHEET in wb.sheetnames:
        ws = wb[SHEET]
    else:
        pos = wb.sheetnames.index('Autocontrolo') + 1 if 'Autocontrolo' in wb.sheetnames \
            else len(wb.sheetnames)
        ws = wb.create_sheet(SHEET, pos)
    if ws['A1'].value is None:
        for j, h in enumerate(HEADERS, 1):
            ws.cell(1, j, h).font = openpyxl.styles.Font(bold=True)
        for col, w in zip('ABC', [34, 22, 70]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A2'

    index = {str(ws.cell(r, 2).value).strip(): r
             for r in range(2, ws.max_row + 1) if ws.cell(r, 2).value}
    atualizadas = adicionadas = 0
    for etar, numero, equip in records:
        row = index.get(str(numero).strip())
        if row:
            atualizadas += 1
        else:
            row = ws.max_row + 1
            index[str(numero).strip()] = row
            adicionadas += 1
        if etar:
            ws.cell(row, 1, etar)
        ws.cell(row, 2, numero)
        ws.cell(row, 3, equip)
        ws.cell(row, 3).alignment = openpyxl.styles.Alignment(wrap_text=True,
                                                              vertical='top')
    try:
        wb.save(master_path)
        return atualizadas, adicionadas, master_path
    except PermissionError:                        # master aberto no Excel
        pending = os.path.splitext(master_path)[0] + '.PENDING.xlsx'
        wb.save(pending)
        return atualizadas, adicionadas, pending


def main(folder, master, incluir_todas=False):
    if not os.path.isdir(folder):
        sys.exit(f'Pasta não encontrada: {folder}')
    if not os.path.isfile(master):
        sys.exit(f'Master não encontrado: {master}')

    records, sem_campo = collect(folder)
    orfas = []
    if not incluir_todas:
        conhecidas = licencas_do_master(master)
        orfas = [n for _, n, _ in records if str(n).strip() not in conhecidas]
        records = [r for r in records if str(r[1]).strip() in conhecidas]
    if not records:
        print('Nada a escrever.')
        return 1
    atualizadas, adicionadas, out = upsert(master, records)
    print(f'{len(records)} licenças lidas de {folder}')
    print(f'{SHEET}: {atualizadas} atualizadas, {adicionadas} adicionadas -> {out}')
    sem_campo = [n for n in sem_campo if any(n == r[1] for r in records)]
    if sem_campo:
        print(f'{len(sem_campo)} sem o campo preenchido: {", ".join(sem_campo[:10])}'
              + (' ...' if len(sem_campo) > 10 else ''))
    if orfas:
        print(f'{len(orfas)} ignoradas por não estarem em Licenses '
              f'(reprocesse o PDF, ou use --all): {", ".join(sorted(orfas)[:10])}'
              + (' ...' if len(orfas) > 10 else ''))
    return 0


if __name__ == '__main__':
    here = os.path.dirname(os.path.abspath(__file__))
    args = [a for a in sys.argv[1:] if a != '--all']
    folder = args[0] if len(args) > 0 else os.path.join(os.path.dirname(here), 'extraction_JSON_LURH')
    master = args[1] if len(args) > 1 else os.path.join(here, 'master_lurh.xlsx')
    sys.exit(main(folder, master, '--all' in sys.argv))
