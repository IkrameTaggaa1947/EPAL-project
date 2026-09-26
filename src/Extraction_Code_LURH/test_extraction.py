"""Regression tests for the LURH extractor.

Fast + deterministic: they run against the frozen '*_extracted.json' outputs
(the "golden" data) rather than re-parsing PDFs, so they finish instantly and
catch any change that breaks the sheet-building or QA logic.

Run either way:
    python test_extraction.py       # plain asserts, no framework
    pytest test_extraction.py
"""
import os
import sys
import glob
import json

# --- Python portatil: garantir que esta pasta esta no sys.path ---------------
# O runtime\python (Python embutivel) corre em modo ISOLADO por causa do
# ficheiro ._pth e, nesse modo, o Python NAO acrescenta a pasta do script ao
# sys.path como faz o Python normal. Sem isto, importar um modulo vizinho
# falha com ModuleNotFoundError. Com um Python instalado, e inofensivo.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import extract_lurh as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
# Extracted JSON now lives in src/extraction_JSON_LURH (sibling of this code folder)
GOLDEN = sorted(glob.glob(os.path.join(os.path.dirname(HERE), 'extraction_JSON_LURH', '*_extracted.json')))


def _load():
    return [json.load(open(p, encoding='utf-8')) for p in GOLDEN]


def test_golden_files_exist():
    assert GOLDEN, "no *_extracted.json golden files found — run the extractor first"


def test_sheets_are_well_formed():
    """Every licence yields all 6 sheets, one Licenses row, and the Nº Licença key
    on every fact row (Power BI relationships depend on this)."""
    for d in _load():
        sheets = E.build_sheets(d)
        assert set(sheets) == set(E.MASTER_SHEETS), sheets.keys()
        assert len(sheets['Licenses']) == 1, d['file_name']
        # one Conditions row per discharge parameter
        assert len(sheets['Conditions']) == len(d['condicoes_descarga']), d['file_name']
        # full autocontrolo and meio-recetor carried through
        assert len(sheets['Autocontrolo']) == len(d['autocontrolo'].get('rows', [])), d['file_name']
        assert len(sheets['MeioRecetor']) == len(d['monitorizacao'].get('rows', [])), d['file_name']
        for name in ('Conditions', 'Autocontrolo', 'MeioRecetor', 'Legislacao', 'Avaliacao'):
            for row in sheets[name]:
                assert row.get('Nº Licença'), (d['file_name'], name)
        assert sheets['Licenses'][0]['TUA/LURH'] == 'LURH', d['file_name']


# Ficheiros com problemas de extracao JA CONHECIDOS e aceites, na tabela de
# monitorizacao do meio recetor (celulas fundidas no PDF de origem). Nao afetam
# os VLE nem o autocontrolo, que alimentam o dashboard. Retirar daqui quando o
# PDF for reextraido corretamente.
MONITORIZACAO_COM_PROBLEMA = {'Proença-a-Nova_L015577.2021.RH5A (09.09.2021 a 08.09.2026).pdf'}


def test_rows_have_key_fields():
    """Nenhuma linha de parametro pode perder as colunas essenciais.

    Recolhe TODAS as falhas antes de rebentar, para se ver o quadro completo em
    vez de so o primeiro ficheiro com problema.
    """
    falhas = []
    for d in _load():
        nome = d['file_name']
        for c in d['condicoes_descarga']:
            if not c.get('Parâmetro'):
                falhas.append(f"{nome}: condicao sem Parâmetro")
            if not (c.get('VLE') or c.get('VLE (% mín. remoção)')):
                falhas.append(f"{nome}: condicao sem VLE ({c.get('Parâmetro')})")
        for a in d['autocontrolo'].get('rows', []):
            if not (a.get('Parâmetro') and a.get('Frequência de amostragem')
                    and a.get('Tipo de amostragem')):
                falhas.append(f"{nome}: autocontrolo incompleto ({a.get('Parâmetro')})")
        if nome in MONITORIZACAO_COM_PROBLEMA:
            continue
        for m in d['monitorizacao'].get('rows', []):
            if not (m.get('Parâmetro') and m.get('Frequência')):
                falhas.append(f"{nome}: monitorizacao incompleta ({m.get('Parâmetro')})")
    assert not falhas, f"{len(falhas)} linha(s) incompleta(s):\n  " + "\n  ".join(falhas[:20])


def test_qa_runs_and_flags_known_issues():
    verdicts = {}
    for d in _load():
        v, reasons = qa.check(d)
        assert v in (qa.OK, qa.FAILED), (d['file_name'], v)
        assert v != qa.FAILED, (d['file_name'], reasons)
        verdicts[d['file_name']] = (v, reasons)

    # NOTA (2026-08-23): o veredito REVIEW foi removido a pedido — qa.check devolve
    # OK ou FAILED, e todos os outros reparos passaram a ser NOTAS informativas.
    # O que continua a importar nao e o veredito, e que a nota seja levantada.

    # Aguas: as datas do documento diferem mesmo das do nome do ficheiro.
    aguas = [k for k in verdicts if k.startswith('Aguas')]
    if aguas:
        v, reasons = verdicts[aguas[0]]
        assert any('does not match filename' in r for r in reasons), (v, reasons)

    # Castanheira: ambiguidade da tabela de 4 colunas com '% remoção'.
    cast = [k for k in verdicts if k.startswith('Castanheira')]
    if cast:
        v, reasons = verdicts[cast[0]]
        assert any('remoção' in r for r in reasons), (v, reasons)


def test_filename_ground_truth_parser():
    """The three filename quirks seen in the corpus must all parse."""
    assert qa.from_filename('Belmonte_L005347_2021_RH5A_12_07_2021_a_11_07_2026.pdf') == {
        'Nº Licença': 'L005347.2021.RH5A',
        'Data de Início': '12-07-2021', 'Data de Validade': '11-07-2026'}
    assert qa.from_filename('Alcorrego__Avis_L014144_2019_RH5A02_09_2019_a_01_09_2024.pdf') == {
        'Nº Licença': 'L014144.2019.RH5A',
        'Data de Início': '02-09-2019', 'Data de Validade': '01-09-2024'}
    assert qa.from_filename('Aldeia_da_Serra_Lic_L003886_2014_RH7_27032014_a_27032021.pdf') == {
        'Nº Licença': 'L003886.2014.RH7',
        'Data de Início': '27-03-2014', 'Data de Validade': '27-03-2021'}


if __name__ == '__main__':
    tests = [v for k, v in sorted(globals().items()) if k.startswith('test_')]
    falhas = 0
    for t in tests:
        try:
            t()
            print(f"PASS  {t.__name__}")
        except Exception as exc:                 # nao parar no 1.o erro. Apanha tudo,
            falhas += 1                          # nao so AssertionError - uma excecao
            kind = type(exc).__name__            # inesperada e uma FALHA, nao um motivo
            if isinstance(exc, AssertionError):  # para abortar a serie.
                kind = ''
            print(f"FALHA {t.__name__}\n        {kind} {str(exc)[:300]}".rstrip())
    print(f"\n{len(tests) - falhas}/{len(tests)} testes OK em {len(GOLDEN)} licencas.")
    sys.exit(1 if falhas else 0)
