# -*- coding: utf-8 -*-
r"""Onde ficam os ficheiros de registo (.log).

PORQUE NAO DENTRO DA PASTA PARTILHADA
-------------------------------------
Os registos ficavam em src\<pipeline>\data\logs\, ou seja, DENTRO da pasta
sincronizada. Isso da dois problemas, e ja deu os dois:

  * quem tem a pasta partilhada apenas para LEITURA nao consegue abrir o
    ficheiro do dia, e rebenta logo no arranque -- antes de processar licenca
    nenhuma:
        PermissionError: [Errno 13] ...\data\logs\watch_20260825.log

  * dois computadores a escrever o mesmo ficheiro do dia levam o OneDrive a
    criar copias em conflito ("watch_20260825-PC-DA-ANA.log"), e o registo
    fica partido aos bocados, sem ninguem dar por isso.

Um registo diz o que aconteceu NESTA maquina: e informacao local, nao e do
projeto. Por isso vai para %LOCALAPPDATA%\EPAL\logs, que nunca sincroniza --
ao lado do epal.local.ini, que ja la vivia pela mesma razao.

Quem quiser juntar os registos de toda a gente num sitio so define EPAL_LOG_DIR.
"""
from __future__ import annotations

import logging
import os
import sys
from datetime import datetime


def log_dir() -> str:
    """A pasta dos registos. Local a este PC, salvo se EPAL_LOG_DIR mandar."""
    v = os.getenv('EPAL_LOG_DIR', '').strip()
    if v:
        return v
    base = os.getenv('LOCALAPPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'EPAL', 'logs')


def _avisar(msg: str) -> None:
    # Sob pythonw.exe (tarefa agendada, sem janela) nao ha consola nenhuma:
    # stderr E stdout sao None. Escrever para None rebentava aqui mesmo.
    for saida in (sys.stderr, sys.stdout):
        if saida is not None:
            print(msg, file=saida)
            return


def add_file_handler(log: logging.Logger, prefixo: str,
                     fmt: logging.Formatter) -> str:
    """Liga o ficheiro do dia ao logger. Devolve o caminho, ou '' se nao deu.

    Nao levanta excecao: ficar sem registo e mau, mas nao e razao para o
    computador se recusar a processar licencas."""
    destino = log_dir()
    try:
        os.makedirs(destino, exist_ok=True)
        caminho = os.path.join(
            destino, '%s_%s.log' % (prefixo, datetime.now().strftime('%Y%m%d')))
        fh = logging.FileHandler(caminho, encoding='utf-8')
    except OSError as exc:
        _avisar('AVISO: nao consegui escrever o registo em %s (%s).\n'
                '       O processamento continua, mas sem ficheiro de registo.'
                % (destino, exc))
        return ''
    fh.setFormatter(fmt)
    log.addHandler(fh)
    return caminho


# ------------------------------------------------------------------ self-check
def _demo():
    import io
    import tempfile

    antigo = os.environ.get('EPAL_LOG_DIR')
    lg = logging.getLogger('logdir-demo')

    def largar():
        # No Windows um handler so removido continua a segurar o ficheiro:
        # tem de fechar, senao a pasta temporaria nao se apaga.
        for h in list(lg.handlers):
            h.close()
        lg.handlers.clear()

    try:
        with tempfile.TemporaryDirectory() as tmp:
            os.environ['EPAL_LOG_DIR'] = tmp
            assert log_dir() == tmp, log_dir()

            largar()
            lg.setLevel(logging.INFO)
            p = add_file_handler(lg, 'demo', logging.Formatter('%(message)s'))
            assert p.startswith(tmp) and os.path.isfile(p), p
            lg.warning('ola')
            for h in lg.handlers:
                h.flush()
            assert 'ola' in io.open(p, encoding='utf-8').read()

            # Um destino impossivel NAO pode rebentar: fica so sem ficheiro.
            # (p e um ficheiro, portanto nao pode ser pasta-mae de nada.)
            largar()
            os.environ['EPAL_LOG_DIR'] = os.path.join(p, 'nao-da')
            assert add_file_handler(lg, 'demo',
                                    logging.Formatter('%(message)s')) == ''
            assert lg.handlers == [], lg.handlers

        # Por omissao: local a maquina, e NUNCA dentro da pasta do projeto.
        os.environ.pop('EPAL_LOG_DIR', None)
        d = log_dir()
        raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        assert d.endswith('logs') and 'EPAL' in d, d
        assert os.path.normcase(raiz) not in os.path.normcase(d), d
    finally:
        largar()
        os.environ.pop('EPAL_LOG_DIR', None)
        if antigo is not None:
            os.environ['EPAL_LOG_DIR'] = antigo
    print('logdir self-check OK')


if __name__ == '__main__':
    _demo()
