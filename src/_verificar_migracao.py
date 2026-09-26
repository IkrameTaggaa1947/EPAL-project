# -*- coding: utf-8 -*-
r"""Verifica se este computador esta pronto para a pasta vinda do SharePoint.

CORRER DEPOIS DE MIGRAR, em cada PC.

O que confirma, por ordem de perigo:

  1. COMPRIMENTO DOS CAMINHOS. E o unico problema que falha em SILENCIO. O
     Windows corta em 260 caracteres; o caminho mais longo dentro do projeto
     tem 146 (uma pasta de visual do Power BI, com nome gerado). Sobram
     113 caracteres para a pasta onde o SharePoint sincroniza. Um caminho de
     equipa e quase sempre MAIS COMPRIDO do que o do OneDrive pessoal:

        C:\Users\<nome>\<Organizacao>\<Site> - <Biblioteca>\EPAL-project

     Se passar dos 113, alguns ficheiros simplesmente nao sincronizam. Nao ha
     erro visivel - o ficheiro nao esta la. Solucao: nomes CURTOS para o site
     e para a biblioteca.

  2. .git NAO PODE SINCRONIZAR. O git escreve ficheiros de bloqueio e um
     indice binario constantemente; um cliente de sincronizacao a mexer nisso
     ao mesmo tempo corrompe o repositorio. Ver docs/MIGRACAO_SHAREPOINT.md.

  3. O Python portatil tem de estar fixado ("manter sempre neste
     dispositivo") e completo - depois de migrar sao ficheiros NOVOS, logo
     voltam a ser marcadores na nuvem.

  4. A ligacao C:\EPAL\powerbi tem de apontar para a NOVA pasta.

Codigo de saida: 0 = pronto, 1 = ha problemas a resolver.
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

MAX_PATH = 260
LIMITE_RAIZ = 113          # 260 - 1 - 146 (o caminho relativo mais longo)


def _titulo(t):
    print()
    print(t)
    print('-' * len(t))


def verificar_comprimentos():
    _titulo('1. Comprimento dos caminhos')
    raiz = len(ROOT)
    print(f'   pasta do projeto : {raiz} caracteres')
    print(f'   maximo seguro    : {LIMITE_RAIZ}')
    demasiado = []
    for dp, dn, fs in os.walk(ROOT):
        if '.git' in dp.split(os.sep):
            continue
        for f in fs:
            p = os.path.join(dp, f)
            if len(p) >= MAX_PATH:
                demasiado.append(p)
    if raiz > LIMITE_RAIZ:
        print(f'   [PROBLEMA] a pasta esta {raiz - LIMITE_RAIZ} caracteres acima do limite.')
        print('              Volte a sincronizar com nomes mais curtos para o site')
        print('              e para a biblioteca no SharePoint.')
    else:
        print(f'   [OK] sobram {LIMITE_RAIZ - raiz} caracteres de margem.')
    if demasiado:
        print(f'   [PROBLEMA] {len(demasiado)} ficheiro(s) ja passam dos {MAX_PATH} caracteres:')
        for p in demasiado[:5]:
            print(f'              {p[-90:]}')
    return raiz <= LIMITE_RAIZ and not demasiado


def verificar_git():
    _titulo('2. Repositorio git')
    g = os.path.join(ROOT, '.git')
    if not os.path.exists(g):
        print('   [OK] nao ha .git dentro da pasta sincronizada.')
        return True
    n = sum(len(fs) for _, _, fs in os.walk(g))
    print(f'   [AVISO] .git existe aqui dentro ({n} ficheiros).')
    print('           Se esta pasta e sincronizada pelo SharePoint, o repositorio')
    print('           pode corromper-se. Ver docs/MIGRACAO_SHAREPOINT.md para as')
    print('           duas saidas possiveis (remoto proprio, ou clone fora da pasta).')
    return True                      # aviso, nao impede de trabalhar


def verificar_python():
    _titulo('3. Python portatil')
    exe = os.path.join(ROOT, 'runtime', 'python', 'python.exe')
    if not os.path.isfile(exe):
        print('   [PROBLEMA] runtime\\python\\python.exe nao existe.')
        print('              Corra Automation files/INSTALAR_PYTHON_PORTATIL.bat uma vez, com internet.')
        return False
    total = fixados = 0
    for dp, _dn, fs in os.walk(os.path.join(ROOT, 'runtime')):
        for f in fs:
            total += 1
            try:
                if os.stat(os.path.join(dp, f)).st_file_attributes & 0x00080000:
                    fixados += 1
            except (OSError, AttributeError):
                pass
    print(f'   ficheiros no runtime : {total}')
    print(f'   fixados no disco     : {fixados}')
    ok = True
    if total and fixados < total:
        print('   [PROBLEMA] o OneDrive pode esvaziar estes ficheiros e partir o Python.')
        print('              Corra Automation files/CONFIGURAR_ESTE_PC.bat (passo 2 fixa-os).')
        ok = False
    else:
        print('   [OK] fixados - o OneDrive nao os esvazia.')
    return ok


def verificar_junction():
    _titulo('4. Ligacao do Power BI')
    sys.path.insert(0, HERE)
    try:
        import epal_config as cfg
    except Exception as exc:
        print(f'   [PROBLEMA] nao consegui ler a configuracao: {exc}')
        return False
    link = os.getenv('EPAL_POWERBI_LINK', r'C:\EPAL\powerbi')
    alvo = cfg.DATA_POWERBI
    if not os.path.isdir(link):
        print(f'   [PROBLEMA] {link} nao existe. Corra Automation files/CONFIGURAR_ESTE_PC.bat.')
        return False
    real = os.path.realpath(link)
    if os.path.normcase(real) != os.path.normcase(os.path.realpath(alvo)):
        print(f'   [PROBLEMA] {link} aponta para a pasta ANTIGA:')
        print(f'              {real}')
        print(f'              devia apontar para: {alvo}')
        print('              Corra Automation files/CONFIGURAR_ESTE_PC.bat para a recriar.')
        return False
    print(f'   [OK] {link} -> {alvo}')
    return True


def verificar_escrita():
    _titulo('5. Permissao de escrita na pasta partilhada')
    # Uma pasta partilhada pode chegar apenas para LEITURA -- e o caso normal
    # quando alguem partilha "Pode ver" em vez de "Pode editar". Ate aqui nada
    # denunciava isso: parecia tudo bem, e depois o primeiro ficheiro que o
    # sistema tentava escrever rebentava com um traceback de Python que nao
    # dizia nada a quem o via:
    #     PermissionError: [Errno 13] ...\data\logs\watch_20260825.log
    # Mais vale perguntar primeiro, em vez de descobrir a meio de uma extracao.
    falhas = []
    for rel in ('src/Extraction_Code_TUA', 'src/Extraction_Code_LURH',
                'src/extraction_JSON_TUA', 'src/extraction_JSON_LURH',
                'src/automation_of_extraction/data',
                'data/pdfs/NEW/TUA', 'data/pdfs/NEW/LURH',
                'data/powerbi'):
        pasta = os.path.join(ROOT, *rel.split('/'))
        if not os.path.isdir(pasta):
            continue
        teste = os.path.join(pasta, '.epal_teste_escrita.tmp')
        try:
            with open(teste, 'w', encoding='utf-8') as fh:
                fh.write('x')
            os.remove(teste)
        except OSError as exc:
            falhas.append((rel, type(exc).__name__))
    if not falhas:
        print('   [OK] consigo escrever em todas as pastas necessarias.')
        return True
    print('   [PROBLEMA] esta pasta esta apenas para LEITURA neste computador.')
    for rel, erro in falhas:
        print(f'              {rel}  ({erro})')
    print()
    print('              NENHUMA alteracao ao codigo resolve isto: quem partilhou')
    print('              a pasta tem de dar permissao de EDICAO a este utilizador')
    print('              (OneDrive/SharePoint: Gerir acesso -> Pode editar).')
    return False


def main() -> int:
    print('=' * 70)
    print(' EPAL - este computador esta pronto para a pasta do SharePoint?')
    print('=' * 70)
    print(f'  pasta: {ROOT}')
    resultados = [verificar_comprimentos(), verificar_git(),
                  verificar_python(), verificar_junction(),
                  verificar_escrita()]
    print()
    print('=' * 70)
    if all(resultados):
        print(' TUDO PRONTO.')
        return 0
    print(' HA PROBLEMAS A RESOLVER - ver [PROBLEMA] acima.')
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
