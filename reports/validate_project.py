"""Validate a PBIP project before opening it in Power BI Desktop.

Catches the classes of error that make Power BI refuse to load the project with
a raw parser message ("Erreur de format TMDL / InvalidLineType"), which is
otherwise only discoverable by opening the file.

Checks
------
TMDL
  1. `///` descriptions must attach to an object declaration — a description
     followed by a blank line raises InvalidLineType: Empty.
  2. Indentation must be tabs, never spaces.
  3. ``` code fences must be balanced and their body indented past the property.
  4. `/* */` block comments are flagged (use `//`).
  5. Partitions must not be truncated (M expression must close).
Modelo
  6. Every relationship column exists.
  7. `ref table` entries in model.tmdl match the files in tables/.
  8. Linguistic bindings in cultures/*.tmdl point at objects that still exist.
  8b. The DataFolder parameter points at a folder that exists (it is an absolute
     path, so it is machine-specific and goes stale whenever the project moves).
Relatório
  9. Every JSON parses.
 10. pages.json order/active page match the page folders on disk.
 11. Every field referenced by a visual exists in the semantic model.

Run:  python validate_project.py [projeto]      (default: EPAL_Licencas_Dashboard)
Exit code 0 = clean, 1 = problems found.
"""
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
# 'queryGroup' é escrito pelo Power BI quando agrupa consultas — por exemplo as
# consultas de erro que ele próprio gera depois de uma atualização falhada. É
# uma declaração legítima: sem ela nesta lista, a descrição /// que a acompanha
# era dada como órfã e o projeto reprovava sem razão.
OBJ = r'^(table|column|measure|relationship|partition|model|hierarchy|level|expression|cultureInfo|ref|annotation|database|queryGroup)\b'


def model_objects(mdir):
    objs = {}
    tdir = os.path.join(mdir, 'tables')
    for fn in os.listdir(tdir):
        if not fn.endswith('.tmdl'):
            continue
        t = open(os.path.join(tdir, fn), encoding='utf-8').read()
        m = re.search(r'^table\s+(\S+)', t, re.M)
        if not m:
            continue
        objs[m.group(1).strip("'")] = {
            (x.group(2) or x.group(1)).strip()
            for x in re.finditer(r"^\t(?:column|measure)\s+('([^']+)'|[^\s=]+)", t, re.M)}
    return objs


def check_tmdl(mdir, problems):
    files = []
    for root, _, fs in os.walk(mdir):
        files += [os.path.join(root, f) for f in fs if f.endswith('.tmdl')]

    for p in sorted(files):
        rel = os.path.relpath(p, mdir)
        L = open(p, encoding='utf-8').read().split('\n')
        text = '\n'.join(L)
        check_orphan_annotations(rel, text, problems)
        check_excel_source(rel, text, problems)

        in_fence = False
        for i, l in enumerate(L, 1):
            s = l.strip()
            # dentro de uma expressão ``` … ``` o texto é DAX/M, não TMDL
            if s == '```' or (l.rstrip().endswith('```') and s != '```'):
                in_fence = not in_fence
                continue
            if in_fence:
                continue
            # O parser TMDL do Power BI NÃO aceita comentários: nem //, nem ///
            # soltos, nem /* */. Só descrições /// coladas a uma declaração.
            # Documentar no README, não no TMDL.
            if s.startswith('///'):
                nxt = L[i].strip() if i < len(L) else ''
                if nxt == '':
                    problems.append(f'{rel}:{i} — descrição /// seguida de linha VAZIA '
                                    f'(Power BI: InvalidLineType: Empty)')
                elif not (nxt.startswith('///') or re.match(OBJ, nxt)):
                    problems.append(f'{rel}:{i} — descrição /// não ligada a uma declaração')
            elif s.startswith('//'):
                problems.append(f'{rel}:{i} — comentário // (Power BI: InvalidLineType: Other). '
                                f'O TMDL não aceita comentários soltos; documentar no README.')
            if l.startswith(' '):
                problems.append(f'{rel}:{i} — indentação com espaços (TMDL exige tabulações)')
            if s.startswith('/*') or s.startswith('*/'):
                problems.append(f'{rel}:{i} — comentário /* */ não é suportado')

        if text.count('```') % 2:
            problems.append(f'{rel} — code fences ``` desequilibradas')

        for i, l in enumerate(L):
            if l.rstrip().endswith('```') and l.strip() != '```':
                depth = len(l) - len(l.lstrip('\t'))
                j = i + 1
                while j < len(L) and L[j].strip() != '```':
                    if L[j].strip() and (len(L[j]) - len(L[j].lstrip('\t'))) <= depth:
                        problems.append(f'{rel}:{j+1} — corpo da expressão sem indentação suficiente')
                    j += 1

        if re.search(r'partition\s', text):
            tail = text.rstrip().split('\n')[-1].strip()
            if re.search(r'=\s*$|,\s*$|\{\{"[^"]*$', tail):
                problems.append(f'{rel} — a partição parece truncada a meio da expressão M')


def check_model(mdir, objs, problems):
    relp = os.path.join(mdir, 'relationships.tmdl')
    if os.path.exists(relp):
        rel = open(relp, encoding='utf-8').read()
        for a, b in re.findall(r'fromColumn: ([^\n]+)\n\ttoColumn: ([^\n]+)', rel):
            for side in (a, b):
                tbl, c = side.split('.', 1)
                if c.strip().strip("'") not in objs.get(tbl.strip(), set()):
                    problems.append(f'relationships.tmdl — {side.strip()} não existe no modelo')

    refs = re.findall(r'^ref table (\S+)', open(os.path.join(mdir, 'model.tmdl'),
                                                encoding='utf-8').read(), re.M)
    missing = set(refs) - set(objs)
    extra = set(objs) - set(refs)
    if missing:
        problems.append(f'model.tmdl — ref table sem ficheiro: {sorted(missing)}')
    if extra:
        problems.append(f'model.tmdl — tabela sem ref table: {sorted(extra)}')

    cdir = os.path.join(mdir, 'cultures')
    if os.path.isdir(cdir):
        for fn in os.listdir(cdir):
            raw = open(os.path.join(cdir, fn), encoding='utf-8').read()
            m = re.search(r'(?:\t{3}.*\n)+', raw)
            if not m:
                continue
            try:
                js = json.loads('\n'.join(l[3:] for l in m.group(0).split('\n') if l.strip()))
            except json.JSONDecodeError as e:
                problems.append(f'cultures/{fn} — linguisticMetadata não é JSON válido: {e}')
                continue
            for k, v in js.get('Entities', {}).items():
                b = v.get('Binding') or v.get('Definition', {}).get('Binding', {})
                ent, prop = b.get('ConceptualEntity'), b.get('ConceptualProperty')
                if prop is None:
                    # an entity may bind to the TABLE itself (no column) — that is
                    # valid; only a reference to a deleted table is a problem
                    if ent not in objs:
                        problems.append(f"cultures/{fn} — '{k}' aponta para a tabela "
                                        f'{ent}, que já não existe')
                elif prop not in objs.get(ent, set()):
                    problems.append(f"cultures/{fn} — '{k}' aponta para "
                                    f'{ent}[{prop}], que já não existe')


# Vocabulário que o próprio Power BI escreve nos visual.json. Foi extraído dos
# visuais gerados por ele neste projeto — ou seja, é uma FOTOGRAFIA de uma
# versão do Desktop, não a especificação. Cada versão nova acrescenta chaves:
# 'howCreated' e 'syncGroup' apareceram depois desta lista ter sido feita e
# produziram 17 alarmes falsos. Por isso uma chave desconhecida passou a ser
# AVISO, não erro: o validador continua a apontar a novidade sem reprovar um
# projeto que o Power BI aceita perfeitamente.
VISUAL_KEYS = {'$schema', 'name', 'position', 'visual', 'filterConfig',
               # escrita pelo Desktop quando um visual nasce de "Sugestões"/copiar
               'howCreated'}
VISUAL_INNER_KEYS = {'visualType', 'query', 'objects', 'visualContainerObjects',
                     'drillFilterOtherVisuals',
                     # written by Desktop on matrices to remember which row
                     # groups are collapsed — legitimate, do not flag it
                     'expansionStates',
                     # grupo de sincronização de segmentações entre páginas
                     'syncGroup'}
QUERY_KEYS = {'queryState', 'sortDefinition', 'isDrillDisabled'}
POSITION_KEYS = {'x', 'y', 'z', 'height', 'width', 'tabOrder', 'angle'}


def check_excel_source(path, text, problems):
    """The partition names both a FILE and a SHEET inside it. When they drift
    apart Power BI fails on load with "La clé ne correspondait à aucune ligne
    dans la table" — an error that says nothing about the real cause."""
    mf = re.search(r'File\.Contents\(DataFolder & "\\([^"]+)\.xlsx"', text)
    mi = re.search(r'Source\{\[Item = "([^"]+)"', text)
    if not mf or not mi:
        return
    feed = os.getenv('EPAL_POWERBI_OUT') or os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'powerbi')
    xlsx = os.path.join(feed, mf.group(1) + '.xlsx')
    if not os.path.exists(xlsx):
        problems.append(f'{path} — o ficheiro {mf.group(1)}.xlsx não existe em {feed}')
        return
    try:
        import openpyxl
        sheets = openpyxl.load_workbook(xlsx, read_only=True).sheetnames
    except Exception as e:
        problems.append(f'{path} — não foi possível ler {mf.group(1)}.xlsx: {e}')
        return
    if mi.group(1) not in sheets:
        problems.append(f'{path} — a partição procura a folha "{mi.group(1)}" mas '
                        f'{mf.group(1)}.xlsx tem {sheets}: o Power BI vai recusar '
                        f'carregar esta tabela.')


def check_datafolder(mdir, problems):
    """The DataFolder parameter is an ABSOLUTE path — Power Query has no notion of
    a path relative to the .pbip, so the value committed to git is whatever the
    last person's machine used. On any other machine (or after the project folder
    moves) the refresh fails with a Power Query error that names a file, never the
    parameter. Flag the drift here, where it is cheap to see.

    The check is deliberately loose: it only compares the TAIL of the path
    ('data\\powerbi' / 'data\\powerbi\\TUA') against this repository's own layout,
    so a teammate with a different user folder is not nagged — only a value that
    points at a folder that does not exist under any plausible root.
    """
    path = os.path.join(mdir, 'expressions.tmdl')
    if not os.path.isfile(path):
        return
    with open(path, encoding='utf-8') as fh:
        m = re.search(r'expression DataFolder = "([^"]+)"', fh.read())
    if not m:
        return
    valor = m.group(1)

    # Caminho estavel (C:\EPAL\powerbi): uma junction criada por
    # fix_powerbi_datafolder.py para o valor ser IGUAL em todos os PCs. Se
    # existe e da na pasta certa deste repositorio, esta correto — nao tem
    # (nem deve ter) uma pasta "data" no nome.
    raiz_repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    feed = os.path.join(raiz_repo, 'data', 'powerbi')
    if os.path.isdir(valor) and os.path.isdir(feed):
        if os.path.normcase(os.path.realpath(valor)) == os.path.normcase(os.path.realpath(feed)):
            return
    cauda = valor.replace('/', '\\').split('\\')
    try:                                    # tudo a partir de 'data' (inclusive)
        i = [c.lower() for c in cauda].index('data')
        relativo = os.path.join(*cauda[i:])
    except ValueError:
        problems.append(f'expressions.tmdl — DataFolder = "{valor}" não passa por '
                        f'uma pasta "data": depois da reorganização o feed vive em '
                        f'<raiz>\\data\\powerbi. Corrigir em Transformar dados -> '
                        f'Gerir parâmetros.')
        return
    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    esperado = os.path.join(raiz, relativo)
    if not os.path.isdir(esperado):
        problems.append(f'expressions.tmdl — DataFolder aponta para "{relativo}", '
                        f'que não existe neste repositório ({esperado}).')
    elif not os.path.isdir(valor) and os.name == 'nt':
        problems.append(f'expressions.tmdl — DataFolder = "{valor}" não existe nesta '
                        f'máquina (é o caminho de quem fez o último commit). '
                        f'Editar para "{esperado}" em Transformar dados -> '
                        f'Gerir parâmetros.')


def check_orphan_annotations(path, text, problems):
    """Two annotations in a row mean a column block was deleted but its trailing
    annotation stayed. Power BI then refuses the project with
    "Impossible de fusionner les objets TMDL … la même propriété : value"."""
    runs = re.findall(r'annotation SummarizationSetBy = Automatic\r?\n\r?\n'
                      r'\t*annotation SummarizationSetBy', text)
    if runs:
        problems.append(f'{path} — {len(runs)} anotação(ões) órfã(s) '
                        f'SummarizationSetBy (bloco de coluna apagado sem a sua '
                        f'anotação): o Power BI recusa o projeto.')


def check_visual_shape(page, vname, d, warnings):
    """Chaves que não conhecemos vão para os AVISOS, não para os problemas.

    A lista de chaves permitidas é uma fotografia de uma versão do Desktop (ver
    VISUAL_KEYS). Tratá-la como verdade absoluta fez o validador reprovar um
    projeto perfeitamente válido 17 vezes — e um validador que grita sem razão
    deixa de ser lido."""
    for level, keys, allowed in (
            ('', d.keys(), VISUAL_KEYS),
            ('/visual', d.get('visual', {}).keys(), VISUAL_INNER_KEYS),
            ('/visual/query', d.get('visual', {}).get('query', {}).keys(), QUERY_KEYS),
            ('/position', d.get('position', {}).keys(), POSITION_KEYS)):
        for k in set(keys) - allowed:
            warnings.append(f'{page}/{vname} — propriedade «{k}» em {level or "/"} '
                            f'não está na lista conhecida. Se o Power BI abrir o '
                            f'projeto, é uma chave nova do Desktop: acrescente-a a '
                            f'{level or "VISUAL_KEYS"} em validate_project.py.')


def check_report(rdir, objs, problems, warnings):
    for root, _, fs in os.walk(rdir):
        for fn in fs:
            if fn.endswith(('.json', '.pbir')) or fn == '.platform':
                p = os.path.join(root, fn)
                try:
                    json.load(open(p, encoding='utf-8'))
                except Exception as e:
                    problems.append(f'{os.path.relpath(p, rdir)} — JSON inválido: {e}')

    pdir = os.path.join(rdir, 'definition', 'pages')
    if not os.path.isfile(os.path.join(pdir, 'pages.json')):
        # legacy single-file report (report.json) — EPAL_TUA_Dashboard still uses
        # it. Nothing below applies; the JSON parse above already covered it.
        return
    pages = json.load(open(os.path.join(pdir, 'pages.json'), encoding='utf-8'))
    dirs = sorted(d for d in os.listdir(pdir) if os.path.isdir(os.path.join(pdir, d)))
    if sorted(pages['pageOrder']) != dirs:
        problems.append(f"pages.json — pageOrder {pages['pageOrder']} != pastas {dirs}")
    if pages['activePageName'] not in dirs:
        problems.append(f"pages.json — activePageName '{pages['activePageName']}' não existe")

    for page in dirs:
        vdir = os.path.join(pdir, page, 'visuals')
        if not os.path.isdir(vdir):
            continue
        for v in sorted(os.listdir(vdir)):
            f = os.path.join(vdir, v, 'visual.json')
            if not os.path.exists(f):
                continue
            blob = open(f, encoding='utf-8').read()
            check_visual_shape(page, v, json.loads(blob), warnings)
            for entity, prop in re.findall(
                    r'"Entity":\s*"([^"]+)"\s*\}\s*\}\s*,\s*"Property":\s*"([^"]+)"', blob):
                if prop not in objs.get(entity, set()):
                    problems.append(f'{page}/{v} — {entity}[{prop}] não existe no modelo')


def main(argv):
    proj = argv[0] if argv else 'EPAL_Licencas_Dashboard'
    base = os.path.join(HERE, proj)
    if not os.path.isdir(base):
        print(f'projeto não encontrado: {base}')
        return 1
    name = next(d for d in os.listdir(base) if d.endswith('.SemanticModel'))
    mdir = os.path.join(base, name, 'definition')
    rdir = os.path.join(base, name.replace('.SemanticModel', '.Report'))

    problems, warnings = [], []
    objs = model_objects(mdir)
    check_tmdl(mdir, problems)
    check_datafolder(mdir, problems)
    check_model(mdir, objs, problems)
    check_report(rdir, objs, problems, warnings)

    print(f'projeto: {proj}')
    print(f'  tabelas: {len(objs)} | objetos: {sum(len(v) for v in objs.values())}')
    if warnings:
        print(f'  AVISOS ({len(warnings)}) — nao impedem o projeto de abrir:')
        for w in warnings:
            print('   ·', w)
    if problems:
        print(f'  PROBLEMAS ({len(problems)}):')
        for p in problems:
            print('   -', p)
        return 1
    print('  tudo válido — o projeto deve abrir no Power BI Desktop')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
