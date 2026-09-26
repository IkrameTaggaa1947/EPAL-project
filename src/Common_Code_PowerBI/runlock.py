# -*- coding: utf-8 -*-
"""Impedir que duas pessoas processem licencas ao mesmo tempo.

PORQUE EXISTE
-------------
Ate agora a pasta era o OneDrive de UMA pessoa, por isso "duas execucoes ao
mesmo tempo" era teorico. Numa biblioteca de equipa deixa de ser: o
PROCESSAR_AGORA.bat existe precisamente para ser clicado por operadores, e nada
impedia dois deles de o clicarem no mesmo minuto.

O que acontece nesse caso: os dois processos leem o master, os dois
reescrevem-no inteiro, e o ultimo a gravar apaga as licencas do outro. Nao ha
erro, nao ha aviso - as linhas desaparecem.

O vigia automatico (watch_all.py) ja tinha um cadeado; as execucoes MANUAIS
nao tinham nenhum. Este modulo da-lhes um.

LIMITES (honestos)
------------------
Isto NAO e um cadeado distribuido. O OneDrive/SharePoint nao da nenhuma
primitiva atomica entre maquinas: dois PCs podem criar o ficheiro no mesmo
instante e so um sobrevive a sincronizacao. E uma convencao de heartbeat que
falha para o lado seguro - "nao corras e explica porque" em vez de "corre e
arrisca". Cobre o caso real (pessoas a trabalhar com minutos de diferenca),
nao uma corrida de milissegundos.
"""
from __future__ import annotations

import json
import os
import platform
import time
import uuid
from datetime import datetime, timezone

# Ao fim de quantos minutos sem sinal se considera que o dono morreu.
STALE_MINUTES = float(os.getenv('EPAL_RUN_LOCK_STALE_MINUTES', '30'))

_OWNER = f"{platform.node() or os.getenv('COMPUTERNAME') or 'desconhecido'}:{uuid.getnode():x}"


def _now():
    return datetime.now(timezone.utc)


def _read(path):
    try:
        with open(path, 'r', encoding='utf-8') as fh:
            d = json.load(fh)
        if isinstance(d, dict) and 'owner' in d and 'heartbeat' in d:
            return d
    except (OSError, ValueError):
        pass
    return None


def _write(path, extra=''):
    payload = {'owner': _OWNER, 'pid': os.getpid(), 'user': os.getenv('USERNAME', '?'),
               'what': extra, 'heartbeat': _now().isoformat()}
    tmp = f"{path}.tmp{os.getpid()}"
    for _ in range(5):
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(tmp, 'w', encoding='utf-8') as fh:
                json.dump(payload, fh)
            os.replace(tmp, path)
            return True
        except OSError:
            time.sleep(1)
    try:
        os.remove(tmp)
    except OSError:
        pass
    return False


def acquire(path, what=''):
    """Tenta ficar com o cadeado. Devolve (conseguiu, mensagem)."""
    if os.getenv('EPAL_RUN_LOCK_FORCE') == '1':
        _write(path, what)
        return True, 'EPAL_RUN_LOCK_FORCE=1 - cadeado ignorado a pedido.'

    actual = _read(path)
    if actual is None:
        if os.path.isfile(path):        # existe mas ilegivel (a sincronizar?)
            return False, ('Ha um ficheiro de cadeado que nao consigo ler '
                           f'({os.path.basename(path)}). Por seguranca nao corro. '
                           'Tente daqui a um minuto.')
        return _write(path, what), ''

    if actual['owner'] == _OWNER and actual.get('pid') == os.getpid():
        return _write(path, what), ''   # ja e nosso: refrescar

    try:
        idade = (_now() - datetime.fromisoformat(actual['heartbeat'])).total_seconds() / 60.0
    except ValueError:
        return False, 'O cadeado tem uma data ilegivel. Por seguranca nao corro.'

    if idade < STALE_MINUTES:
        quem = actual.get('user') or actual['owner']
        mesma = actual['owner'] == _OWNER
        onde = 'NESTE computador' if mesma else f"no computador de {quem}"
        return False, (f"Ja esta a correr {onde} (ha {idade:.0f} min). "
                       f"NAO corri, para nao estragar o master. "
                       f"Espere que termine e tente outra vez.")

    return _write(path, what), (f"O cadeado de {actual.get('user') or actual['owner']} "
                                f"parece abandonado (ha {idade:.0f} min sem sinal) - assumi-o.")


def release(path):
    """Larga o cadeado, se for nosso."""
    actual = _read(path)
    if actual and actual['owner'] == _OWNER and actual.get('pid') == os.getpid():
        try:
            os.remove(path)
        except OSError:
            pass


class held:
    """`with runlock.held(path, 'TUA') as (ok, msg):` - larga sempre no fim."""

    def __init__(self, path, what=''):
        self.path, self.what, self.ok, self.msg = path, what, False, ''

    def __enter__(self):
        self.ok, self.msg = acquire(self.path, self.what)
        return self

    def __exit__(self, *exc):
        if self.ok:
            release(self.path)
        return False


# ------------------------------------------------------------------ self-check
def _demo():
    import tempfile
    global _OWNER
    with tempfile.TemporaryDirectory() as tmp:
        p = os.path.join(tmp, 'run.lock')

        with held(p, 'TUA') as h:
            assert h.ok, h.msg
            assert os.path.isfile(p)

            # outra maquina, com o cadeado fresco -> tem de recusar
            meu = _OWNER
            _OWNER = 'outro-pc:abc'
            ok, msg = acquire(p)
            assert not ok, 'devia ter recusado'
            assert 'Ja esta a correr' in msg, msg
            _OWNER = meu
        assert not os.path.isfile(p), 'o cadeado devia ter sido largado no fim'

        # cadeado velho -> pode ser assumido
        _write(p)
        d = _read(p)
        d['owner'] = 'pc-morto:1'
        d['heartbeat'] = datetime(2020, 1, 1, tzinfo=timezone.utc).isoformat()
        with open(p, 'w', encoding='utf-8') as fh:
            json.dump(d, fh)
        ok, msg = acquire(p)
        assert ok and 'abandonado' in msg, (ok, msg)
        release(p)
    print('runlock self-check OK')


if __name__ == '__main__':
    _demo()
