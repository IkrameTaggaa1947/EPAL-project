# -*- coding: utf-8 -*-
"""Prepare everything that has to be e-mailed — Power Automate only sends it.

Business logic stays in Python; the flow in Power Automate reads one table and
sends one message per row. That keeps the flow trivial (a beginner can read it)
and every decision testable here.

Produces:
    outputs/renewal/Pedido Renovação_<ETAR>_<Nº>.xlsx   one per licence to renew
    outputs/alerts/Alertas.xlsx                         one row = one e-mail

Alert types
    Renovação    licence expiring inside RENEWAL_WINDOW_MONTHS — the filled-in
                 renewal form goes with it as an attachment
    Extração     the QA gate found broken data (truncated parameters, bad dates)
    Dados        fields the renewal form needs and nobody has filled yet

Recipients come from data/reference/contactos.xlsx (created empty on first run).
Rows whose recipient is still unknown fall back to FALLBACK_EMAIL, so nothing is
silently dropped.

Run:  python src/make_alerts.py
"""
import datetime
import os
import re
import subprocess
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
ROOT = os.path.dirname(HERE)

FEED = os.getenv('EPAL_POWERBI_OUT', os.path.join(ROOT, 'data', 'powerbi'))
REFERENCE = os.path.join(ROOT, 'data', 'reference')
RENEWAL_DIR = os.path.join(ROOT, 'outputs', 'renewal')
ALERTS_DIR = os.path.join(ROOT, 'outputs', 'alerts')
TEMPLATE = os.path.join(REFERENCE, 'Pedido Renovação LURH_ETAR_XXX_v1.xlsx')
CONTACTS = os.path.join(REFERENCE, 'contactos.xlsx')

FALLBACK_EMAIL = 'licencas@epal.pt'      # change to the real mailbox
MISSING = 'Em falta'


def read(path, sheet=None):
    import openpyxl
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    it = ws.values
    try:
        hdr = [str(c).strip() if c is not None else '' for c in next(it)]
    except StopIteration:
        wb.close()
        return []
    out = [dict(zip(hdr, r)) for r in it]
    wb.close()
    return out


def sanitize(s):
    return re.sub(r'[^\w\-.]+', '_', str(s)).strip('_')[:60]


# --------------------------------------------------------------------------- #
#  contacts
# --------------------------------------------------------------------------- #
def load_contacts():
    """ETAR -> e-mail. Creates the file with one row per licence on first run."""
    import openpyxl
    rows = read(CONTACTS)
    if rows:
        return {str(r.get('ETAR') or '').strip(): str(r.get('Email') or '').strip()
                for r in rows if str(r.get('Email') or '').strip()}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Contactos'
    ws.append(['ETAR', 'Email', 'Responsável', 'Notas'])
    seen = set()
    for r in read(os.path.join(FEED, 'Licenses.xlsx')):
        etar = str(r.get('Estabelecimento') or '').strip()
        if etar and etar not in seen:
            seen.add(etar)
            ws.append([etar, '', '', ''])
    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 32
    os.makedirs(REFERENCE, exist_ok=True)
    wb.save(CONTACTS)
    print(f'  criado {os.path.relpath(CONTACTS, ROOT)} com {len(seen)} ETAR — '
          f'preencher a coluna Email')
    return {}


# --------------------------------------------------------------------------- #
#  renewal forms, one Excel per licence
# --------------------------------------------------------------------------- #
def build_renewal_files(renovacao):
    """Fill the official template with the values already computed in the feed."""
    import openpyxl
    if not os.path.exists(TEMPLATE):
        print(f'! template não encontrado: {TEMPLATE}', file=sys.stderr)
        return {}
    os.makedirs(RENEWAL_DIR, exist_ok=True)

    by_licence = defaultdict(list)
    for r in renovacao:
        by_licence[str(r.get('Nº TUA') or '')].append(r)

    made = {}
    for tua, fields in by_licence.items():
        values = {str(f.get('Campo') or '').strip(): str(f.get('Valor') or '') for f in fields}
        etar = str(fields[0].get('ETAR') or '')
        validade = str(fields[0].get('Validade') or '')
        regime = str(fields[0].get('Regime') or '')

        wb = openpyxl.load_workbook(TEMPLATE)
        ws = wb['Inf para requerimento']
        for row in range(1, ws.max_row + 1):
            label = ws.cell(row=row, column=2).value
            if label is None:
                continue
            key = str(label).strip()
            # the template labels are longer than our field names, match on both
            hit = values.get(key) or next(
                (v for k, v in values.items() if k and (k in key or key in k)), None)
            if hit is not None:
                ws.cell(row=row, column=3).value = hit
        ws.cell(row=1, column=3).value = f'{etar}  |  {tua}  |  {regime}  |  Validade: {validade}'

        name = f'Pedido Renovação_{sanitize(etar)}_{sanitize(tua)}.xlsx'
        wb.save(os.path.join(RENEWAL_DIR, name))
        made[tua] = name
    print(f'  {len(made)} formulários de renovação em outputs/renewal/')
    return made


# --------------------------------------------------------------------------- #
#  alerts
# --------------------------------------------------------------------------- #
def renewal_alerts(renovacao, files, contacts, today):
    alerts = []
    by_licence = defaultdict(list)
    for r in renovacao:
        by_licence[str(r.get('Nº TUA') or '')].append(r)

    for tua, fields in sorted(by_licence.items(),
                              key=lambda kv: int(kv[1][0].get('Dias p/ expirar') or 0)):
        head = fields[0]
        etar = str(head.get('ETAR') or '')
        dias = int(head.get('Dias p/ expirar') or 0)
        validade = str(head.get('Validade') or '')
        total = len(fields)
        falta = sum(1 for f in fields if str(f.get('Valor') or '').strip() == MISSING)
        preenchidos = total - falta

        # 6 months is the legal minimum to submit; 8 is our preparation window
        urgencia = 'ALTA' if dias <= 183 else 'MÉDIA'
        corpo = (
            f'Caro(a) responsável pela {etar},\n\n'
            f'A licença {tua} caduca a {validade} — faltam {dias} dias.\n'
            f'O pedido de renovação tem de dar entrada com um mínimo de 6 meses '
            f'de antecedência, pelo que é necessário preparar o processo agora.\n\n'
            f'Em anexo segue o formulário de requerimento já preenchido com a '
            f'informação que temos ({preenchidos} de {total} campos).\n'
            f'Faltam {falta} campos, assinalados com "{MISSING}" — precisamos '
            f'que sejam completados e devolvidos.\n\n'
            f'Obrigado,\nGestão de Licenças')
        alerts.append({
            'Tipo': 'Renovação', 'Prioridade': urgencia, 'Licenca': tua, 'ETAR': etar,
            'Validade': validade, 'DiasParaExpirar': dias,
            'Para': contacts.get(etar, FALLBACK_EMAIL),
            'Assunto': f'[Renovação {urgencia}] {etar} — licença caduca a {validade} '
                       f'({dias} dias)',
            'Corpo': corpo,
            'Anexo': f'outputs/renewal/{files[tua]}' if tua in files else '',
            'Enviado': 'Não',
        })
    return alerts


def quality_alerts(today):
    """Run the QA gate and turn its findings into one e-mail for the data owner."""
    try:
        res = subprocess.run([sys.executable, os.path.join(HERE, 'qa_feed.py')],
                             capture_output=True, text=True, timeout=300)
    except Exception as exc:
        print(f'! QA não pôde correr: {exc}', file=sys.stderr)
        return []
    out = res.stdout
    errors = [l.strip(' -') for l in out.splitlines()
              if l.startswith('   - ') and 'ERROS' not in l]
    if res.returncode == 0 and 'ERROS' not in out:
        return []
    bloco = '\n'.join(f'  • {e}' for e in errors) or '  (ver saída do QA)'
    return [{
        'Tipo': 'Extração', 'Prioridade': 'ALTA', 'Licenca': '', 'ETAR': '',
        'Validade': '', 'DiasParaExpirar': '',
        'Para': FALLBACK_EMAIL,
        'Assunto': f'[Extração] Problemas detetados no feed de {today:%d/%m/%Y}',
        'Corpo': ('O controlo de qualidade encontrou problemas nos dados que '
                  'alimentam o dashboard:\n\n' + bloco +
                  '\n\nEstes registos chegam ao dashboard como estão. '
                  'É preciso rever a extração das licenças em causa.\n\n'
                  'Gestão de Licenças'),
        'Anexo': '', 'Enviado': 'Não',
    }]


def missing_data_alert(renovacao, today):
    """One summary of the fields nobody has been able to fill, across licences."""
    falta = defaultdict(int)
    licences = set()
    for r in renovacao:
        if str(r.get('Valor') or '').strip() == MISSING:
            falta[str(r.get('Campo') or '')] += 1
            licences.add(str(r.get('Nº TUA') or ''))
    if not falta:
        return []
    linhas = '\n'.join(f'  • {c} — em falta em {n} licença(s)'
                       for c, n in sorted(falta.items(), key=lambda kv: -kv[1])[:15])
    return [{
        'Tipo': 'Dados', 'Prioridade': 'MÉDIA', 'Licenca': '', 'ETAR': '',
        'Validade': '', 'DiasParaExpirar': '',
        'Para': FALLBACK_EMAIL,
        'Assunto': f'[Dados em falta] {len(falta)} campos impedem fechar '
                   f'{len(licences)} pedidos de renovação',
        'Corpo': ('Os campos abaixo são pedidos no requerimento de renovação e '
                  'não existem em nenhuma das nossas fontes:\n\n' + linhas +
                  '\n\nSem esta informação os pedidos não podem ser submetidos. '
                  'Agradecemos indicação de onde a obter.\n\nGestão de Licenças'),
        'Anexo': '', 'Enviado': 'Não',
    }]


def write_alerts(alerts, today):
    """One sheet formatted as a real Excel Table — Power Automate needs a table."""
    import openpyxl
    from openpyxl.worksheet.table import Table, TableStyleInfo
    os.makedirs(ALERTS_DIR, exist_ok=True)
    cols = ['Id', 'Data', 'Tipo', 'Prioridade', 'Licenca', 'ETAR', 'Validade',
            'DiasParaExpirar', 'Para', 'Assunto', 'Corpo', 'Anexo', 'Enviado']
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Alertas'
    ws.append(cols)
    for i, a in enumerate(alerts, 1):
        a['Id'] = f'{today:%Y%m%d}-{i:03d}'
        a['Data'] = f'{today:%d/%m/%Y}'
        ws.append([str(a.get(c, '')) for c in cols])
    last = ws.max_row if ws.max_row > 1 else 2
    table = Table(displayName='Alertas', ref=f'A1:M{last}')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True)
    ws.add_table(table)
    for col, width in zip('ABCDEFGHIJKLM',
                          (14, 12, 12, 11, 20, 30, 12, 8, 28, 60, 15, 45, 10)):
        ws.column_dimensions[col].width = width
    path = os.path.join(ALERTS_DIR, 'Alertas.xlsx')
    wb.save(path)
    return path


def main():
    today = datetime.date.today()
    print('make_alerts:', today.strftime('%d/%m/%Y'))
    renovacao = read(os.path.join(FEED, 'Renovacao_Proposta.xlsx'))
    if not renovacao:
        print('! Renovacao_Proposta.xlsx vazio — correr make_powerbi_extras.py primeiro',
              file=sys.stderr)
        return 1

    contacts = load_contacts()
    files = build_renewal_files(renovacao)

    alerts = (renewal_alerts(renovacao, files, contacts, today)
              + quality_alerts(today)
              + missing_data_alert(renovacao, today))
    path = write_alerts(alerts, today)

    por_tipo = defaultdict(int)
    for a in alerts:
        por_tipo[a['Tipo']] += 1
    sem_contacto = sum(1 for a in alerts if a['Para'] == FALLBACK_EMAIL and a['ETAR'])
    print(f'  {len(alerts)} alertas -> {os.path.relpath(path, ROOT)}')
    for t, n in sorted(por_tipo.items()):
        print(f'     {t}: {n}')
    if sem_contacto:
        print(f'  ! {sem_contacto} ETAR sem e-mail em data/reference/contactos.xlsx '
              f'(usam {FALLBACK_EMAIL})')
    return 0


if __name__ == '__main__':
    sys.exit(main())
