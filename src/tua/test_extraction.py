"""Regression tests for the TUA extractor.

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

import extract_tua as E
import qa

HERE = os.path.dirname(os.path.abspath(__file__))
GOLDEN = sorted(glob.glob(os.path.join(HERE, '*_extracted.json')))


def _load():
    return [json.load(open(p, encoding='utf-8')) for p in GOLDEN]


def test_golden_files_exist():
    assert GOLDEN, "no *_extracted.json golden files found — run the extractor first"


def test_sheets_are_well_formed():
    """Every certificate yields all 5 sheets, one Licenses row, and the Nº TUA key
    on every fact row (Power BI relationships depend on this)."""
    for d in _load():
        sheets = E.build_sheets(d)
        assert set(sheets) == set(E.MASTER_SHEETS), sheets.keys()
        assert len(sheets['Licenses']) == 1, d['file_name']
        # one Conditions row per discharge parameter
        assert len(sheets['Conditions']) == len(d['condicoes_rejeicao']), d['file_name']
        # full autocontrolo carried through (the fix: Entrada / monitor-only params)
        assert len(sheets['Autocontrolo']) == len(d['autocontrolo']), d['file_name']
        for name in ('Conditions', 'Autocontrolo', 'Legislacao', 'Avaliacao'):
            for row in sheets[name]:
                assert row.get('Nº TUA'), (d['file_name'], name)


def test_qa_runs_and_flags_known_issues():
    verdicts = {}
    for d in _load():
        v, reasons = qa.check(d)
        assert v in (qa.OK, qa.REVIEW, qa.FAILED)
        verdicts[d['file_name']] = (v, reasons)

    # Aldeia da Ribeira lost its whole dates section -> must be flagged, not OK
    aldeia = [k for k in verdicts if k.startswith('Aldeia')]
    if aldeia:
        v, reasons = verdicts[aldeia[0]]
        assert v == qa.REVIEW, (v, reasons)
        assert any('Validade' in r for r in reasons)


def test_per_file_workbook_is_readable():
    """The per-certificate Excel opens, starts with the Resumo identity card, and
    Resumo still carries every single-value field (so the layout never silently
    drops extracted information). Section sheets are added when 'sections' is
    present (live extraction); golden JSON baselines predate that field."""
    import tempfile
    import openpyxl
    for d in _load():
        with tempfile.TemporaryDirectory() as tmp:
            _, xp = E.write_outputs(d, tmp)
            wb = openpyxl.load_workbook(xp)
            assert wb.sheetnames[0] == 'Resumo', (d['file_name'], wb.sheetnames)
            assert len(wb.sheetnames) == 1 + len(d.get('sections', [])), (d['file_name'], wb.sheetnames)
            blob = E.norm('\n'.join(
                str(v) for ws in wb.worksheets for row in ws.iter_rows(values_only=True)
                for v in row if v not in (None, '')))
            for sec in ('dados_gerais', 'enquadramento', 'localizacao', 'caracterizacao'):
                for k, v in d[sec].items():
                    assert not v or E.norm(v) in blob, (d['file_name'], sec, k, v)


def test_all_sections_extracted_if_pdf_available():
    """When a TUA PDF is present locally, the generic pass yields the full set of
    sections (>=10) — each with a clean header and >=1 row. Skips in a bare
    checkout (PDFs are git-ignored)."""
    import glob
    pdfs = glob.glob(os.path.join(HERE, '**', '*TUA*.pdf'), recursive=True)
    if not pdfs:
        return
    d = E.extract(sorted(pdfs)[0])
    assert len(d['sections']) >= 10, [s['code'] for s in d['sections']]
    codes = {s['code'] for s in d['sections']}
    assert {'3.13', '3.16', '3.19', '3.20', '3.21'} <= codes, sorted(codes)
    for s in d['sections']:
        assert s['columns'] and s['rows'], s['code']


def test_filename_cross_check():
    fn = qa.from_filename('Foios_TUA20230905002577_04.02.2024_a_03.02.2029.pdf')
    assert fn['Nº TUA'] == 'TUA20230905002577'
    assert fn['Data de Validade'] == '03-02-2029'
    assert qa.from_filename('not_a_tua_file.pdf') is None


if __name__ == '__main__':
    qa._demo()
    tests = [v for k, v in sorted(globals().items()) if k.startswith('test_')]
    for t in tests:
        t()
        print(f"PASS  {t.__name__}")
    print(f"\nAll {len(tests)} tests passed on {len(GOLDEN)} golden certificates.")
