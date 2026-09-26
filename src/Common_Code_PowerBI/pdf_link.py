# -*- coding: utf-8 -*-
r"""Uma ligação clicável, por licença, para o PDF original.

PARA QUE SERVE
--------------
Permite pôr no dashboard um botão "Abrir licença" que abre o PDF da licença
que está a ser vista. O Power BI não consegue abrir um ficheiro a partir de
um caminho qualquer — precisa de um URL numa coluna marcada como *Web URL*.
Este módulo constrói essa coluna.

DOIS CENÁRIOS, E É POR ISSO QUE O ENDEREÇO BASE É CONFIGURÁVEL
--------------------------------------------------------------
1. **Depois da mudança para o SharePoint** (o caso bom): cada PDF tem um
   endereço https:// real, igual para toda a gente, e o botão funciona tanto
   no Power BI Desktop como no Power BI Service (relatório publicado).
   Define-se UMA vez, para toda a equipa:

       EPAL_PDF_BASE_URL=https://<tenant>.sharepoint.com/sites/<site>/... \
                         /Licences%20project/EPAL-project/data/pdfs

2. **Antes da mudança** (ou sem configuração): cai para um `file:///` com o
   caminho desta máquina. Abre no Power BI **Desktop**, mas não no Service, e
   é específico de quem gerou o feed — serve para experimentar, não para
   distribuir.

O URL é sempre construído a partir do sítio REAL onde o PDF está
(data/pdfs/NEW/TUA, OLD/LURH, ...), procurado pelo nome do ficheiro: o
`file_name` que vem do master nem sempre é o nome actual do PDF no arquivo.
Quando não há PDF correspondente, a coluna fica VAZIA — e o botão deve
desligar-se sozinho, em vez de abrir um endereço partido.
"""
from __future__ import annotations

import glob
import os
import pathlib
import urllib.parse

COLUNA = 'Ficheiro da Licença'

# ===================================================================
#  DEPOIS DA MUDANÇA PARA O SHAREPOINT: trocar a linha do BASE_PADRAO
#  (mais abaixo) por        BASE_PADRAO = BASE_SHAREPOINT
#  e voltar a correr        make_powerbi_all.py
#  É só isso. Confirmar primeiro o endereço abaixo: abrir um PDF qualquer
#  no SharePoint pelo browser e comparar com o que está aqui — sobretudo
#  se a pasta 'EPAL-project' existe mesmo lá dentro.
# ===================================================================
BASE_SHAREPOINT = ('https://privepal.sharepoint.com/sites/2026-InternshipatEPAL'
                   '/Documentos%20Partilhados/General/Deliverables'
                   '/Licences%20project/EPAL-project/data/pdfs')

# Servidor local (Automation files/SERVIDOR_PDFS.bat). É preciso porque o Power BI NÃO abre
# ficheiros locais a partir de um botão: um endereço file:/// é ignorado, e
# só http:// e https:// funcionam. Enquanto o servidor estiver a correr, o
# botão funciona nesta máquina.
BASE_LOCALHOST = 'http://localhost:%s' % os.getenv('EPAL_PDF_PORT', '8000')

BASE_PADRAO = BASE_LOCALHOST

# Subpastas do arquivo, pela ordem em que se prefere encontrar o ficheiro:
# uma licença que exista nas duas fica com a cópia actual (NEW).
SUBPASTAS = [os.path.join('NEW', 'TUA'), os.path.join('NEW', 'LURH'),
             os.path.join('OLD', 'TUA'), os.path.join('OLD', 'LURH')]


def base_url() -> str:
    """O endereço base a usar. Precedência: ambiente > .ini > BASE_PADRAO.

    Pôr '-' em qualquer um dos dois primeiros força o modo file:/// local,
    para quem quiser trabalhar sem rede."""
    v = os.getenv('EPAL_PDF_BASE_URL', '').strip()
    if not v:
        try:
            import epal_config as cfg
            v = cfg.get('pdf_base_url', section='powerbi').strip()
        except Exception:
            v = ''
    if v == '-':
        return ''
    return (v or BASE_PADRAO).rstrip('/')


def indexar(pasta_pdfs: str) -> dict:
    """{nome_do_ficheiro: caminho_relativo_com_barras} para todo o arquivo."""
    idx = {}
    for sub in SUBPASTAS:
        for p in glob.glob(os.path.join(pasta_pdfs, sub, '*.pdf')):
            nome = os.path.basename(p)
            idx.setdefault(nome, f"{sub.replace(os.sep, '/')}/{nome}")
    return idx


def url_para(nome_ficheiro: str, indice: dict, pasta_pdfs: str, base: str) -> str:
    """URL clicável para este PDF, ou '' se o ficheiro não estiver no arquivo."""
    rel = indice.get(nome_ficheiro)
    if not rel:
        return ''
    if base:
        # cada segmento é codificado à parte para as barras se manterem barras
        return base + '/' + '/'.join(urllib.parse.quote(s) for s in rel.split('/'))
    absoluto = os.path.join(pasta_pdfs, *rel.split('/'))
    return pathlib.Path(absoluto).as_uri()        # file:/// já codificado


def acrescentar_coluna(rows, pasta_pdfs: str, campo_nome: str = 'file_name'):
    """Acrescenta COLUNA a cada linha. Devolve (com_ligacao, sem_ligacao)."""
    indice = indexar(pasta_pdfs)
    base = base_url()
    com = sem = 0
    for r in rows:
        u = url_para(str(r.get(campo_nome, '')), indice, pasta_pdfs, base)
        r[COLUNA] = u
        if u:
            com += 1
        else:
            sem += 1
    return com, sem


# ------------------------------------------------------------------ self-check
def _demo():
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        alvo = os.path.join(tmp, 'NEW', 'TUA')
        os.makedirs(alvo)
        nome = 'Foios TUA20230905002577 (04.02.2024 a 03.02.2029).pdf'
        open(os.path.join(alvo, nome), 'wb').close()
        idx = indexar(tmp)
        assert nome in idx, idx

        # 1. com endereco base (SharePoint): espacos e parenteses codificados
        u = url_para(nome, idx, tmp, 'https://x.sharepoint.com/sites/s/pdfs')
        assert u.startswith('https://x.sharepoint.com/sites/s/pdfs/NEW/TUA/'), u
        assert ' ' not in u and '%20' in u, u
        assert '/NEW/TUA/' in u, 'as barras da pasta nao podem ser codificadas'

        # 2. sem endereco base: file:/// desta maquina
        u2 = url_para(nome, idx, tmp, '')
        assert u2.startswith('file:///'), u2

        # 3. por omissao segue o BASE_PADRAO (hoje '' = file:/// local)
        os.environ.pop('EPAL_PDF_BASE_URL', None)
        assert base_url() == BASE_PADRAO, base_url()
        os.environ['EPAL_PDF_BASE_URL'] = '-'      # '-' forca o modo local
        assert base_url() == '', base_url()
        # o endereco do SharePoint continua guardado, pronto a usar
        assert BASE_SHAREPOINT.startswith('https://'), BASE_SHAREPOINT
        u3 = url_para(nome, idx, tmp, BASE_SHAREPOINT)
        assert u3.startswith(BASE_SHAREPOINT + '/NEW/TUA/'), u3
        os.environ['EPAL_PDF_BASE_URL'] = 'https://outro/sitio/'
        assert base_url() == 'https://outro/sitio', base_url()
        os.environ.pop('EPAL_PDF_BASE_URL', None)

        # 4. ficheiro que nao existe -> vazio, para o botao se desligar
        assert url_para('nao_existe.pdf', idx, tmp, 'https://x/y') == ''

        # 5. a coluna e acrescentada a todas as linhas
        rows = [{'file_name': nome}, {'file_name': 'nao_existe.pdf'}]
        com, sem = acrescentar_coluna(rows, tmp)
        assert (com, sem) == (1, 1), (com, sem)
        assert rows[0][COLUNA] and rows[1][COLUNA] == ''
    print('pdf_link self-check OK')


if __name__ == '__main__':
    _demo()
