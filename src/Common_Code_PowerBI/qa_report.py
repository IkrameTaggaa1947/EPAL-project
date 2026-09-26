# -*- coding: utf-8 -*-
"""Escrever o qa_report.xlsx sem perder historico e sem crescer para sempre.

O QUE ESTAVA MAL
----------------
Os dois pipelines tinham a MESMA funcao, com dois problemas:

  1. A leitura do relatorio existente estava dentro de um
     `except Exception: pass`. Se o ficheiro estivesse corrompido ou a
     meio de uma sincronizacao, o historico TODO era descartado em
     silencio e substituido apenas pelas linhas desta execucao. Um
     registo de auditoria que se apaga sozinho e pior do que nao existir.

  2. Nunca era truncado nem rodado. Cresce a cada licenca, para sempre.

O QUE FAZ AGORA
---------------
  * Se o relatorio existente NAO puder ser lido, NAO o sobrepoe: escreve
    as linhas novas ao lado, num '<nome>.PENDING.xlsx', e diz porque.
  * Se o relatorio estiver aberto no Excel, mesma coisa (era o que ja
    fazia) - mas agora esse PENDING volta a ser integrado na execucao
    seguinte, em vez de ficar esquecido.
  * Quando passa de MAX_ROWS linhas, a versao completa e arquivada em
    data/logs/qa_report_<data>.xlsx e o ficheiro de trabalho fica so com
    as mais recentes. Nada e apagado - so movido.
"""
from __future__ import annotations

import os
from datetime import datetime

import pandas as pd

# Linhas que o ficheiro de trabalho guarda antes de rodar. 2000 chega para
# mais de uma dezena de execucoes completas do corpus.
MAX_ROWS = int(os.getenv('EPAL_QA_REPORT_MAX_ROWS', '2000'))


def pending_path(path: str) -> str:
    return os.path.splitext(path)[0] + '.PENDING.xlsx'


def _read(path: str):
    """(DataFrame, erro). erro != None significa 'existe mas nao consegui ler'."""
    if not os.path.isfile(path):
        return None, None
    try:
        return pd.read_excel(path), None
    except Exception as exc:
        return None, exc


def append(rows, path, archive_dir=None, max_rows=None, log=None):
    """Acrescenta `rows` ao relatorio em `path`. Devolve (total_linhas, mensagem)."""
    if not rows:
        return 0, ''
    max_rows = max_rows or MAX_ROWS
    nome = os.path.basename(path)
    novo = pd.DataFrame(rows)

    # 1. integrar um PENDING deixado por uma execucao anterior bloqueada
    pend = pending_path(path)
    if os.path.isfile(pend):
        antes, erro = _read(pend)
        if antes is not None:
            novo = pd.concat([antes, novo], ignore_index=True)
            os.remove(pend)
            if log:
                log.info(f"  {os.path.basename(pend)} integrado no relatorio e removido.")

    # 2. ler o relatorio actual - e NAO o sobrepor se nao der para ler
    actual, erro = _read(path)
    if erro is not None:
        novo.to_excel(pend, index=False)
        msg = (f"{nome} existe mas NAO pode ser lido ({erro}). O historico NAO foi "
               f"tocado; estas {len(novo)} linhas ficaram em {os.path.basename(pend)}. "
               f"Verifique o ficheiro e volte a correr.")
        if log:
            log.error(f"  {msg}")
        return len(novo), msg
    if actual is not None:
        novo = pd.concat([actual, novo], ignore_index=True)

    # 3. rodar quando fica grande - arquivar tudo, manter as recentes
    msg = ''
    if len(novo) > max_rows:
        destino = archive_dir or os.path.join(os.path.dirname(path), 'data', 'logs')
        try:
            os.makedirs(destino, exist_ok=True)
            arq = os.path.join(destino, f"{os.path.splitext(nome)[0]}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
            novo.to_excel(arq, index=False)
            novo = novo.tail(max_rows).reset_index(drop=True)
            msg = f"{nome} passou de {max_rows} linhas - historico completo arquivado em {arq}."
            if log:
                log.info(f"  {msg}")
        except Exception as exc:                 # arquivar falhou -> NAO truncar
            if log:
                log.warning(f"  Nao consegui arquivar {nome} ({exc}) - fica por truncar.")

    # 4. escrever
    try:
        novo.to_excel(path, index=False)
    except PermissionError:
        novo.to_excel(pend, index=False)
        msg = (f"{nome} esta aberto no Excel - escrevi {os.path.basename(pend)} em vez disso. "
               f"Feche-o; a proxima execucao integra-o sozinha.")
        if log:
            log.warning(f"  {msg}")
    return len(novo), msg


# ------------------------------------------------------------------ self-check
def _demo():
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        p = os.path.join(tmp, 'qa_report.xlsx')

        # 1. primeira escrita
        n, _ = append([{'file_name': 'a.pdf', 'verdict': 'OK'}], p)
        assert n == 1, n

        # 2. acrescenta, nao substitui
        n, _ = append([{'file_name': 'b.pdf', 'verdict': 'FAILED'}], p)
        assert n == 2, n
        assert list(pd.read_excel(p)['file_name']) == ['a.pdf', 'b.pdf']

        # 3. relatorio ilegivel -> historico INTACTO, linhas novas ao lado
        with open(p, 'wb') as fh:
            fh.write(b'isto nao e um xlsx')
        n, msg = append([{'file_name': 'c.pdf', 'verdict': 'OK'}], p)
        assert 'NAO pode ser lido' in msg, msg
        assert os.path.isfile(pending_path(p)), 'as linhas novas deviam ir para o PENDING'
        assert open(p, 'rb').read() == b'isto nao e um xlsx', 'o ficheiro mau foi sobreposto!'

        # 4. reparado -> o PENDING volta a entrar
        os.remove(p)
        n, _ = append([{'file_name': 'd.pdf', 'verdict': 'OK'}], p)
        assert not os.path.isfile(pending_path(p)), 'o PENDING devia ter sido integrado'
        assert list(pd.read_excel(p)['file_name']) == ['c.pdf', 'd.pdf'], pd.read_excel(p)

        # 5. rotacao: arquiva tudo, mantem as recentes
        logs = os.path.join(tmp, 'logs')
        append([{'file_name': f'{i}.pdf', 'verdict': 'OK'} for i in range(20)],
               p, archive_dir=logs, max_rows=5)
        assert len(pd.read_excel(p)) == 5, len(pd.read_excel(p))
        arqs = os.listdir(logs)
        assert len(arqs) == 1, arqs
        assert len(pd.read_excel(os.path.join(logs, arqs[0]))) == 22, 'o arquivo deve ter TUDO'
    print('qa_report self-check OK')


if __name__ == '__main__':
    _demo()
