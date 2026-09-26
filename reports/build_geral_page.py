"""Generate the visuals of the 'Visão Geral' page (section0) of
EPAL_Licencas_Dashboard, in PBIR format.

The page is generated rather than hand-edited so the layout stays consistent
and can be regenerated after model changes. Every visual reads MEASURES from
the Medidas table; the only column used for grouping/legend is
Licenses[Estado da Licença], which is the key into the Dim_Estado dimension.

Run:  python build_geral_page.py            (rebuilds section0/visuals)
      python build_geral_page.py --check    (validates field references only)

Layout, 1280 x 964:
    header 0-58 | left rail x=8 w=232 (slicers) | main area x=248 w=1024
    KPI row y=64 h=96 | mapa + donut y=168 h=280 | ano + região y=456 h=256
    tabela de licenças y=724 h=232 (largura total)
"""
import argparse
import json
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PAGE = os.path.join(HERE, 'EPAL_Licencas_Dashboard', 'EPAL_Licencas_Dashboard.Report',
                    'definition', 'pages', 'section0')
MODEL = os.path.join(HERE, 'EPAL_Licencas_Dashboard', 'EPAL_Licencas_Dashboard.SemanticModel',
                     'definition')

SCHEMA = ('https://developer.microsoft.com/json-schemas/fabric/item/report/definition/'
          'visualContainer/2.10.0/schema.json')

# EPAL palette (same tones as the WWTP report)
AZUL, LARANJA, VERDE, VERMELHO, CINZA = '#0D6ABF', '#FF911B', '#1E8449', '#C0392B', '#95A5A6'
PANEL, PANEL2, FUNDO = '#C6E4E0', '#D0DBE5', '#EDF4F3'

ESTADO_CORES = {
    'Caducada': VERMELHO,
    'A expirar (< 8 meses)': LARANJA,
    'Válida': VERDE,
    'Sem data': CINZA,
}


# ------------------------------------------------------------- expressions
def lit(v):
    return {'expr': {'Literal': {'Value': v}}}


def txt(s):
    return lit(f"'{s}'")


def num(n):
    return lit(f'{n}D')


def boolean(b):
    return lit('true' if b else 'false')


def solid(color):
    return {'solid': {'color': lit(f"'{color}'")}}


def col(entity, prop):
    return {'Column': {'Expression': {'SourceRef': {'Entity': entity}}, 'Property': prop}}


def measure(prop, entity='Medidas'):
    return {'Measure': {'Expression': {'SourceRef': {'Entity': entity}}, 'Property': prop}}


def avg(entity, prop):
    return {'Aggregation': {'Expression': col(entity, prop)['Column'] and col(entity, prop),
                            'Function': 1}}


def proj(field, entity, prop, native=None, active=False, qref=None):
    p = {'field': field, 'queryRef': qref or f'{entity}.{prop}',
         'nativeQueryRef': native or prop}
    if active:
        p['active'] = True
    return p


def col_proj(entity, prop, active=False, display=None):
    p = proj(col(entity, prop), entity, prop, active=active)
    if display:
        p['displayName'] = display
    return p


def measure_proj(prop, display=None):
    p = proj(measure(prop), 'Medidas', prop)
    if display:
        p['displayName'] = display
    return p


def avg_proj(entity, prop):
    return proj(avg(entity, prop), entity, prop,
                qref=f'Avg({entity}.{prop})')


def datapoint_colors(entity, prop, cores):
    """One fill per category value, so the four states always keep their colour."""
    out = []
    for value, color in cores.items():
        out.append({
            'properties': {'fill': solid(color)},
            'selector': {'data': [{'scopeId': {'Comparison': {
                'ComparisonKind': 0,
                'Left': col(entity, prop),
                'Right': {'Literal': {'Value': f"'{value}'"}},
            }}}]},
        })
    return out


def title(text, color='#1F3B57', size=12, align='left'):
    return {'title': [{'properties': {
        'show': boolean(True), 'text': txt(text),
        'fontColor': solid(color), 'fontSize': num(size), 'alignment': txt(align),
    }}]}


def background(color, transparency=0):
    return {'background': [{'properties': {
        'show': boolean(True), 'color': solid(color), 'transparency': num(transparency)}}]}


def border(color='#B9CBD8'):
    return {'border': [{'properties': {'show': boolean(True), 'color': solid(color),
                                       'radius': num(6)}}]}


def visual(name, x, y, w, h, vtype, query=None, objects=None, container=None,
           z=1000, tab=1000, sort=None):
    v = {'visualType': vtype}
    if query is not None:
        v['query'] = {'queryState': query}
        # sortDefinition is a sibling of queryState INSIDE query — placing it at
        # the visual level makes Power BI reject the file with
        # "Une propriété supplémentaire « sortDefinition » a été incluse".
        if sort is not None:
            v['query']['sortDefinition'] = sort
    if objects:
        v['objects'] = objects
    if container:
        v['visualContainerObjects'] = container
    v['drillFilterOtherVisuals'] = True
    return {
        '$schema': SCHEMA,
        'name': name,
        'position': {'x': x, 'y': y, 'z': z, 'height': h, 'width': w, 'tabOrder': tab},
        'visual': v,
    }


# ------------------------------------------------------------------ visuals
def build():
    V = []

    # --- cabeçalho -------------------------------------------------------
    V.append(visual('faixa_titulo', 0, 0, 1280, 58, 'shape', z=0, tab=-1, objects={
        'shape': [{'properties': {'tileShape': txt('rectangle')}}],
        'fill': [{'properties': {'fillColor': solid(AZUL), 'transparency': num(0)},
                  'selector': {'id': 'default'}}],
        'outline': [{'properties': {'show': boolean(False)}}],
    }))

    V.append(visual('titulo', 16, 8, 620, 42, 'textbox', z=100, tab=0, objects={
        'general': [{'properties': {'paragraphs': [{'textRuns': [{
            'value': 'EPAL — Licenças  ·  Visão Geral',
            'textStyle': {'fontSize': '20pt', 'fontWeight': 'bold', 'color': '#FFFFFF',
                          'fontFamily': 'Segoe UI'},
        }]}]}}],
    }))

    V.append(visual('data_referencia', 980, 12, 288, 36, 'card', z=100, tab=100,
                    query={'Values': {'projections': [measure_proj('Data de Referência')]}},
                    objects={
                        'labels': [{'properties': {'color': solid('#FFFFFF'),
                                                   'fontSize': num(11)}}],
                        'categoryLabels': [{'properties': {'show': boolean(False)}}],
                    }))

    # --- KPIs ------------------------------------------------------------
    kpis = [
        ('kpi_total', 'Total de Licenças', 'Total de licenças', AZUL),
        ('kpi_caducadas', 'Licenças Caducadas', 'Caducadas', VERMELHO),
        ('kpi_expirar', 'Licenças a Expirar', 'A expirar (< 8 meses)', LARANJA),
        ('kpi_validas', 'Licenças Válidas', 'Válidas', VERDE),
        ('kpi_semdata', 'Licenças Sem Data', 'Sem data de validade', CINZA),
    ]
    for i, (name, med, rotulo, cor) in enumerate(kpis):
        V.append(visual(name, 248 + i * 206, 64, 200, 96, 'card', z=2000, tab=200 + i,
                        query={'Values': {'projections': [measure_proj(med)]}},
                        objects={
                            'labels': [{'properties': {'color': solid(cor), 'fontSize': num(28),
                                                       'bold': boolean(True)}}],
                            'categoryLabels': [{'properties': {'show': boolean(False)}}],
                        },
                        container={**title(rotulo, size=11, align='center'),
                                   **background('#FFFFFF'), **border()}))

    # --- filtros (rail esquerdo) ----------------------------------------
    def slicer(name, entity, prop, header, y, h, mode='Basic', tab=300):
        return visual(name, 8, y, 232, h, 'slicer', z=3000, tab=tab,
                      query={'Values': {'projections': [col_proj(entity, prop, active=True)]}},
                      objects={
                          'data': [{'properties': {'mode': txt(mode)}}],
                          'general': [{'properties': {'orientation': lit('2D')}}],
                          'header': [{'properties': {'show': boolean(True), 'text': txt(header),
                                                     'fontColor': solid('#1F3B57'),
                                                     'fontSize': num(10),
                                                     'bold': boolean(True)}}],
                      },
                      container={**background('#FFFFFF'), **border()})

    V.append(slicer('slicer_tipo', 'Licenses', 'TUA/LURH', 'Tipo de licença', 64, 88,
                    tab=300))
    V.append(slicer('slicer_estado', 'Dim_Estado', 'Estado', 'Estado de validade', 160, 128,
                    tab=310))
    V.append(slicer('slicer_regiao', 'Licenses', 'REGIÃO', 'Região', 296, 150, tab=320))
    V.append(slicer('slicer_arh', 'Licenses', 'ARH', 'ARH', 454, 96, mode='Dropdown', tab=330))
    V.append(slicer('slicer_municipio', 'Licenses', 'Município', 'Município', 558, 74,
                    mode='Dropdown', tab=340))

    V.append(visual('nota_criterio', 8, 640, 232, 72, 'textbox', z=3000, tab=350, objects={
        'general': [{'properties': {'paragraphs': [{'textRuns': [{
            'value': ('Classificação: Caducada = validade anterior a hoje · '
                      'A expirar = expira em menos de 8 meses · Sem data = o documento '
                      'não indica validade.'),
            'textStyle': {'fontSize': '8pt', 'color': '#5A6B7B', 'fontFamily': 'Segoe UI'},
        }]}]}}],
    }))

    # --- mapa ------------------------------------------------------------
    V.append(visual('mapa_expiracao', 248, 168, 620, 280, 'azureMap', z=4000, tab=400,
                    query={
                        'Category': {'projections': [col_proj('Licenses', 'Estabelecimento',
                                                              active=True)]},
                        'Series': {'projections': [col_proj('Licenses', 'Estado da Licença')]},
                        'X': {'projections': [avg_proj('Licenses', 'Longitude')]},
                        'Y': {'projections': [avg_proj('Licenses', 'Latitude')]},
                        'Tooltips': {'projections': [col_proj('Licenses', 'Validade (ISO)'),
                                                     col_proj('Licenses', 'REGIÃO')]},
                    },
                    objects={
                        'dataPoint': datapoint_colors('Licenses', 'Estado da Licença',
                                                      ESTADO_CORES),
                        'bubbleLayer': [{'properties': {'show': boolean(True),
                                                        'size': num(6)}}],
                        'legend': [{'properties': {'show': boolean(True),
                                                   'position': txt('BottomCenter')}}],
                    },
                    container={**title('Mapa de validade e expiração'),
                               **background('#FFFFFF'), **border()}))

    # --- donut por estado -------------------------------------------------
    V.append(visual('donut_estado', 876, 168, 396, 280, 'donutChart', z=4000, tab=410,
                    query={
                        'Category': {'projections': [col_proj('Dim_Estado', 'Estado',
                                                              active=True)]},
                        'Y': {'projections': [measure_proj('Licenças por Estado')]},
                    },
                    objects={
                        'dataPoint': datapoint_colors('Dim_Estado', 'Estado', ESTADO_CORES),
                        'legend': [{'properties': {'show': boolean(True),
                                                   'position': txt('BottomCenter'),
                                                   'fontSize': num(9)}}],
                        'labels': [{'properties': {'show': boolean(True),
                                                   'labelStyle': txt('Both'),
                                                   'fontSize': num(9)}}],
                        'slices': [{'properties': {'innerRadiusRatio': num(55)}}],
                    },
                    container={**title('Licenças por estado de validade'),
                               **background('#FFFFFF'), **border()}))

    # NOTA: o modelo foi reduzido às três tabelas do esquema geral (Licenses,
    # Dim_Estado, Medidas). A tabela de datas Calendario foi removida, pelo que
    # o gráfico "Expirações por ano" já não existe. A página é agora mantida à
    # mão; este gerador serve de referência do estilo dos visuais.

    # --- licenças por região ---------------------------------------------
    V.append(visual('bar_regiao', 876, 456, 396, 256, 'barChart', z=4000, tab=430,
                    query={
                        'Category': {'projections': [col_proj('Licenses', 'REGIÃO',
                                                              active=True)]},
                        'Y': {'projections': [measure_proj('Total de Licenças')]},
                    },
                    objects={
                        'dataPoint': [{'properties': {'fill': solid(AZUL)}}],
                        'legend': [{'properties': {'show': boolean(False)}}],
                        'labels': [{'properties': {'show': boolean(True), 'fontSize': num(9)}}],
                        'categoryAxis': [{'properties': {'show': boolean(True),
                                                         'fontSize': num(9)}}],
                    },
                    sort={'sort': [{'field': measure('Total de Licenças'),
                                    'direction': 'Descending'}],
                          'isDefaultSort': True},
                    container={**title('Licenças por região'),
                               **background('#FFFFFF'), **border()}))

    # --- tabela de todas as licenças -------------------------------------
    # Ordenada por data de validade ascendente: o que expira primeiro fica
    # no topo. Acompanha todos os filtros da página.
    V.append(visual('tabela_licencas', 8, 724, 1264, 232, 'tableEx', z=4000, tab=440,
                    query={'Values': {'projections': [
                        col_proj('Licenses', 'Estabelecimento', active=True, display='ETAR'),
                        col_proj('Licenses', 'TUA/LURH', display='Tipo'),
                        col_proj('Licenses', 'Nº TUA', display='Nº do título'),
                        col_proj('Licenses', 'REGIÃO', display='Região'),
                        col_proj('Licenses', 'Município'),
                        col_proj('Licenses', 'Estado da Licença', display='Estado'),
                        col_proj('Licenses', 'Validade (ISO)', display='Validade'),
                        measure_proj('Dias até Expirar', display='Dias até expirar'),
                        col_proj('Licenses', 'Nível de tratamento', display='Tratamento'),
                    ]}},
                    sort={'sort': [{'field': col('Licenses', 'Validade (ISO)'),
                                    'direction': 'Ascending'}],
                          'isDefaultSort': True},
                    objects={
                        'grid': [{'properties': {'gridVertical': boolean(True),
                                                 'gridVerticalColor': solid('#DCE6EC'),
                                                 'rowPadding': num(2)}}],
                        'columnHeaders': [{'properties': {'fontColor': solid('#FFFFFF'),
                                                          'backColor': solid(AZUL),
                                                          'bold': boolean(True),
                                                          'fontSize': num(9)}}],
                        'values': [{'properties': {'fontSize': num(9),
                                                   'backColorSecondary': solid('#F4F8FA')}}],
                    },
                    container={**title('Todas as licenças — ordenadas por data de validade'),
                               **background('#FFFFFF'), **border()}))

    return V


# ------------------------------------------------------------- validation
def model_fields():
    """Every table/column/measure defined in the semantic model."""
    fields, tables = set(), set()
    for fn in os.listdir(os.path.join(MODEL, 'tables')):
        if not fn.endswith('.tmdl'):
            continue
        text = open(os.path.join(MODEL, 'tables', fn), encoding='utf-8').read()
        m = re.search(r'^table\s+(\S+)', text, re.M)
        if not m:
            continue
        table = m.group(1).strip("'")
        tables.add(table)
        for mm in re.finditer(r"^\t(?:column|measure)\s+('([^']+)'|[^\s=]+)", text, re.M):
            fields.add((table, (mm.group(2) or mm.group(1)).strip()))
    return tables, fields


def check(visuals):
    tables, fields = model_fields()
    problems = []
    for v in visuals:
        blob = json.dumps(v, ensure_ascii=False)
        for entity, prop in re.findall(
                r'"Entity":\s*"([^"]+)"\}\},\s*"Property":\s*"([^"]+)"', blob):
            if entity not in tables:
                problems.append(f"{v['name']}: unknown table '{entity}'")
            elif (entity, prop) not in fields:
                problems.append(f"{v['name']}: '{entity}'[{prop}] not in the model")
    return problems


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    a = ap.parse_args(argv)

    visuals = build()
    problems = check(visuals)
    if problems:
        print('FIELD REFERENCE PROBLEMS:')
        for p in sorted(set(problems)):
            print('  -', p)
        if not a.check:
            return 1
    else:
        print(f'field references OK ({len(visuals)} visuals)')
    if a.check:
        return 0

    vdir = os.path.join(PAGE, 'visuals')
    if os.path.isdir(vdir):
        shutil.rmtree(vdir)
    os.makedirs(vdir)
    for v in visuals:
        d = os.path.join(vdir, v['name'])
        os.makedirs(d)
        with open(os.path.join(d, 'visual.json'), 'w', encoding='utf-8') as f:
            json.dump(v, f, ensure_ascii=False, indent=2)
            f.write('\n')
    print(f'wrote {len(visuals)} visuals to section0/visuals')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
