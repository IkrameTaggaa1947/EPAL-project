"""Regression tests for the LURH extractor.

Fast + deterministic: they run against the frozen '*_extracted.json' outputs
(the "golden" data) rather than re-parsing PDFs, so they finish instantly and
catch any change that breaks the sheet-building or QA logic.

Run either way:
    python test_extraction.py       # plain asserts, no framework
    pytest test_extraction.py
"""
import os
import glob
import json

import extract_lurh as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
GOLDEN = sorted(glob.glob(os.path.join(HERE, '*_extracted.json')))


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


def test_rows_have_key_fields():
    """No parameter row may lose its core columns (the page-break repairs)."""
    for d in _load():
        for c in d['condicoes_descarga']:
            assert c.get('Parâmetro'), (d['file_name'], c)
            assert c.get('VLE') or c.get('VLE (% mín. remoção)'), (d['file_name'], c)
        for a in d['autocontrolo'].get('rows', []):
            assert a.get('Parâmetro') and a.get('Frequência de amostragem') \
                and a.get('Tipo de amostragem'), (d['file_name'], a)
        for m in d['monitorizacao'].get('rows', []):
            assert m.get('Parâmetro') and m.get('Frequência'), (d['file_name'], m)


def test_qa_runs_and_flags_known_issues():
    verdicts = {}
    for d in _load():
        v, reasons = qa.check(d)
        assert v in (qa.OK, qa.REVIEW, qa.FAILED)
        assert v != qa.FAILED, (d['file_name'], reasons)
        verdicts[d['file_name']] = (v, reasons)

    # Aguas: the document's dates genuinely differ from the filename -> REVIEW
    aguas = [k for k in verdicts if k.startswith('Aguas')]
    if aguas:
        v, reasons = verdicts[aguas[0]]
        assert v == qa.REVIEW, (v, reasons)
        assert any('does not match filename' in r for r in reasons)

    # Castanheira: 4-column '% remoção' table ambiguity -> flagged for a human
    cast = [k for k in verdicts if k.startswith('Castanheira')]
    if cast:
        v, reasons = verdicts[cast[0]]
        assert v == qa.REVIEW, (v, reasons)
        assert any('remoção' in r for r in reasons)


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
    test_golden_files_exist()
    test_sheets_are_well_formed()
    test_rows_have_key_fields()
    test_qa_runs_and_flags_known_issues()
    test_filename_ground_truth_parser()
    print(f'all tests OK ({len(GOLDEN)} golden files)')
