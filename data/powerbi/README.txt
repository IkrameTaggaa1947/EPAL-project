POWER BI DATA — all files the dashboards read
=============================================

Licenses.csv / Conditions.csv / Autocontrolo.csv / Legislacao.csv / Avaliacao.csv
    The COMBINED feed (TUA + LURH, 269 licences) for EPAL_Licencas_Dashboard.pbip.
    Regenerate with:  python extraction/make_powerbi_all.py

TUA/  (same table names, TUA only, 129 licences)
    The feed for EPAL_TUA_Dashboard.pbip. Written automatically by the TUA
    pipeline (extraction/tua) every time a certificate is processed.

LURH/  (same table names, LURH only, 134 licences)
    Exported automatically from master_lurh.xlsx (project root) by
    extraction/make_powerbi_all.py. Edit master_lurh.xlsx, then re-run it.

static_tables.xlsx
    Hand-maintained lookups merged into the CSVs at export time
    (region/ARH registry, renewal requests, sampling frequencies).
    Edit here, then re-run the pipeline export or make_powerbi_all.py.

TabelaCentroide_ETAR.xlsx
    ETAR asset registry (464 assets) copied from data/reference/ with one added
    column, Chave_ETAR — the normalised ETAR name that joins it to
    Licenses[Chave_ETAR]. Written by make_powerbi_all.py; edit the source in
    data/reference/, never this copy.
    It carries CODMAXIMO, which is the same code as WWTP_History[CodigoMaximo]:
    that is the bridge between a licence and its laboratory history.
    227/269 licences match by name; the rest have a blank Chave_ETAR — add them
    to OVERRIDES in src/etar_key.py as they are confirmed.

WWTP_History.xlsx
    Laboratory measurement history (434k rows) — source of WWTP_Report.pbix.
    Joins to the licences through TabelaCentroide_ETAR: WWTP_History[CodigoMaximo]
    = TabelaCentroide_ETAR[CODMAXIMO] → [Chave_ETAR] → Licenses[Chave_ETAR].

NOTE: WWTP_Report.pbix still points to the old location of WWTP_History.xlsx.
Re-point it once in Power BI Desktop: Transform data → Data source settings →
Change Source → select this folder's WWTP_History.xlsx.
