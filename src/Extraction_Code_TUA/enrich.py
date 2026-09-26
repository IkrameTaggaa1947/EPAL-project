"""Enrich the extracted sheets with the static lookup tables (static_tables.xlsx),
reproducing the old pipeline's merge so the Power BI dashboard runs on the new data:

  Conditions   += Critério 1..7                     (Criteria, on Parâmetro + Avaliação)
  Licenses     += REGIÃO / ARH / Município / Estado / Tratamento   (AdvT, on Estabelecimento)
                  + Data do pedido de renovação no Siliamb          (Requests, on Código TURH)
  Autocontrolo += Nº análises requeridas             (Frequencies, on Frequência de amostragem)

Kept as a POST step: the raw master_tua.xlsx stays clean (pure extraction); the
powerbi_csv/ feed carries the enriched, dashboard-ready columns. Robust to the
lookup file being open in Excel (then it returns the sheets unchanged).

The Criteria sheet uses abbreviations (pH, CBO5, O&G, SST) with no units, so the
new parameter names must have their '(unit)' suffix stripped before matching —
that trailing text is exactly what the old (truncating) extractor happened to drop.
"""
import os
import re
import pandas as pd

# abbreviation (as in the Criteria sheet) -> full parameter name (as extracted)
ABBR = {'CBO5': 'Carência Bioquímica de Oxigénio',
        'CQO': 'Carência Química de Oxigénio',
        'O&G': 'Óleos e Gorduras',
        'SST': 'Sólidos Suspensos Totais'}


def _norm(s):
    """Parameter key: drop '(unit)' parentheticals, expand abbreviations, strip
    punctuation/case — so 'Carência Bioquímica de Oxigénio (mg/L O2)' and 'CBO5' meet."""
    s = re.sub(r'\s*\([^)]*\)', '', str(s or ''))
    s = ABBR.get(s.strip(), s)
    s = re.sub(r'[^\w\s]', '', s.lower())
    return re.sub(r'\s+', ' ', s).strip()


def _laws(s):
    """A legislation/assessment text -> set of word tokens, for overlap matching."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return set()
    t = re.sub(r'[^\w\s]', ' ', str(s).lower())
    return {x for x in t.split() if len(x) > 2}


def load_static(path):
    try:
        return pd.read_excel(path, sheet_name=None)
    except Exception:
        return None                                   # locked/missing → skip enrichment


def enrich_conditions(cond, criteria):
    crit_cols = [c for c in criteria.columns if str(c).startswith('Critério')]
    lk = criteria.copy()
    lk['_p'] = lk['Parâmetro'].map(_norm)
    lk['_a'] = lk['Avaliação da Conformidade Legal'].map(_laws)
    out = cond.copy()
    for c in crit_cols:
        out[c] = ''
    for i, row in out.iterrows():
        cand = lk[lk['_p'] == _norm(row.get('Parâmetro'))]
        if cand.empty:
            for c in crit_cols:
                out.at[i, c] = 'not found'
            continue
        an = _laws(row.get('Avaliação da Conformidade Legal'))
        ov = cand[cand['_a'].apply(lambda s: bool(s & an))]   # disambiguate by assessment
        use = ov if not ov.empty else cand                    # fall back to parameter-only
        for c in crit_cols:
            vals = pd.unique(use[c].dropna())
            out.at[i, c] = '; '.join(map(str, vals)) if len(vals) else ''
    return out


def enrich_licenses(lic, advt, requests):
    out = lic.copy()
    if advt is not None and 'GNA official' in advt.columns:
        a = advt.copy()
        a['_g'] = a['GNA official'].astype(str).str.strip().str.lower()
        cols = [c for c in ['REGIÃO', 'ARH', 'Município', 'Estado', 'Tratamento implementado']
                if c in a.columns]
        a = a[['_g'] + cols].drop_duplicates('_g')
        out['_g'] = out['Estabelecimento'].astype(str).str.strip().str.lower()
        out = out.merge(a, on='_g', how='left').drop(columns='_g')
    if requests is not None:
        r = requests.copy()
        r.columns = [str(c).strip() for c in r.columns]
        dc = [c for c in r.columns if 'Siliamb' in c]
        if 'Código' in r.columns and dc:
            r = r[['Código', dc[0]]].rename(
                columns={'Código': '_c', dc[0]: 'Data do pedido de renovação no Siliamb'})
            r['_c'] = r['_c'].astype(str).str.replace(r'\s+', '', regex=True)
            r = r.drop_duplicates('_c')
            out['_c'] = out['Código TURH'].astype(str).str.replace(r'\s+', '', regex=True)
            out = out.merge(r, on='_c', how='left').drop(columns='_c')
    return out


def enrich_autocontrolo(auto, freq):
    out = auto.copy()
    if freq is None or 'frequency' not in freq.columns:
        return out
    f = freq.copy()
    f.columns = [str(c).strip() for c in f.columns]
    ncol = [c for c in f.columns if 'lises' in c or 'requerid' in c]
    if ncol:
        f = f[['frequency', ncol[0]]].rename(columns={'frequency': '_f'}).drop_duplicates('_f')
        out['_f'] = out['Frequência de amostragem'].astype(str).str.strip()
        f['_f'] = f['_f'].astype(str).str.strip()
        out = out.merge(f, on='_f', how='left').drop(columns='_f')
    return out


def enrich_all(sheets, static_path):
    """sheets: {name -> DataFrame}. Returns (enriched_sheets, ok)."""
    st = load_static(static_path)
    if st is None:
        return sheets, False
    s = dict(sheets)
    if 'Conditions' in s and 'Criteria' in st:
        s['Conditions'] = enrich_conditions(s['Conditions'], st['Criteria'])
    if 'Licenses' in s:
        s['Licenses'] = enrich_licenses(s['Licenses'], st.get('AdvT'), st.get('Requests'))
    if 'Autocontrolo' in s and 'Frequencies' in st:
        s['Autocontrolo'] = enrich_autocontrolo(s['Autocontrolo'], st['Frequencies'])
    return s, True


def _default_static(master_path):
    return os.getenv('EPAL_STATIC_TABLES',
                     os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(master_path)))), 'data', 'powerbi'), 'static_tables.xlsx'))


if __name__ == '__main__':
    # Standalone: enrich the master's sheets and report what matched.
    here = os.path.dirname(os.path.abspath(__file__))
    master = os.path.join(here, 'master_tua.xlsx')
    sheets = pd.read_excel(master, sheet_name=None)
    enriched, ok = enrich_all(sheets, _default_static(master))
    print('static tables read:', ok)
    for n, df in enriched.items():
        print(f'  {n}: {df.shape[1]} cols')
