# -*- coding: utf-8 -*-
r"""Aponta o dashboard Power BI para a pasta de dados DESTA maquina.

O PROBLEMA
----------
O Power Query nao tem caminhos relativos. O parametro `DataFolder` (em
expressions.tmdl) e um caminho ABSOLUTO. Como a pasta do projeto e
sincronizada para um sitio diferente em cada PC, o caminho que esta gravado
la dentro esta certo para quem gravou e errado para todos os outros - dai o
erro "coluna nao encontrada" / "nao foi possivel encontrar a origem de dados".

A solucao antiga era reescrever o DataFolder com o caminho deste PC. Mas
expressions.tmdl esta DENTRO da pasta partilhada: a correcao sincronizava
para toda a gente e partia o dashboard do colega seguinte, que voltava a
corrigir, e assim sucessivamente. Era o proprio problema que epal_config.py
existe para evitar - definicoes de maquina nunca vivem na pasta partilhada.

A SOLUCAO
---------
Damos a TODAS as maquinas o MESMO caminho absoluto:

    C:\EPAL\powerbi   ->  (junction)  ->  <este PC>\EPAL-project\data\powerbi

A junction e criada por este script (nao precisa de direitos de
administrador). O DataFolder passa a ser sempre "C:\EPAL\powerbi", igual em
todos os PCs, portanto expressions.tmdl deixa de mudar e deixa de haver
ping-pong entre colegas.

Se a junction nao puder ser criada (politica da empresa, C: protegido), o
script avisa e volta ao comportamento antigo - o caminho deste PC - para o
dashboard continuar a funcionar aqui.

IMPORTANTE: corra isto com o Power BI Desktop FECHADO. O Desktop reescreve os
ficheiros do projeto quando grava e desfaz esta alteracao.

Uso:
    python fix_powerbi_datafolder.py            # corrige
    python fix_powerbi_datafolder.py --check    # so verifica, nao altera
    python fix_powerbi_datafolder.py --no-link  # forca o caminho deste PC
"""
from __future__ import annotations

import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import epal_config as cfg                                        # noqa: E402

PATTERN = re.compile(r'(expression\s+DataFolder\s*=\s*")([^"]*)(")')

# Caminhos absolutos que nao sejam o DataFolder sao sempre um bug: qualquer
# tabela deve referir  DataFolder & "\ficheiro.xlsx".
ABS_PATH = re.compile(r'"[A-Za-z]:\\\\?[^"]*"')


# O caminho estavel, igual em todos os PCs. Alteravel para testar ou se C:
# nao servir nesta organizacao.
STABLE_LINK = os.getenv('EPAL_POWERBI_LINK', r'C:\EPAL\powerbi')


def current_value(text: str) -> str | None:
    m = PATTERN.search(text)
    return m.group(2) if m else None


def _same_dir(a: str, b: str) -> bool:
    return os.path.normcase(os.path.realpath(a)) == os.path.normcase(os.path.realpath(b))


def ensure_junction(link: str, target: str) -> tuple[bool, str]:
    """Faz `link` apontar para `target`. Devolve (conseguiu, mensagem).

    Uma junction (mklink /J) NAO precisa de direitos de administrador, ao
    contrario de um symlink. os.rmdir sobre uma junction remove apenas a
    ligacao, nunca o conteudo do destino.
    """
    if not os.path.isdir(target):
        return False, f"a pasta de dados ainda nao existe: {target}"

    if os.path.isdir(link):
        if _same_dir(link, target):
            return True, f"{link} ja aponta para esta maquina."
        try:
            os.rmdir(link)                 # so remove a ligacao / pasta vazia
        except OSError as exc:
            return False, (f"{link} existe e nao e uma ligacao para este projeto "
                           f"(nao consegui remover: {exc}). Apague-a a mao se nao "
                           f"tiver nada la dentro.")
    elif os.path.isfile(link):
        return False, f"{link} existe e e um ficheiro, nao uma pasta."
    elif os.path.lexists(link):
        # JUNCTION PARTIDA: o destino deixou de existir (a pasta do projeto foi
        # movida ou mudou de nome). Nesse estado nao e isdir NEM exists -- so
        # lexists a ve -- portanto os dois ramos acima nao lhe tocavam e o
        # mklink seguinte rebentava com "ja existe", deixando a ligacao partida
        # para sempre. E exatamente o caso que a reparacao devia resolver.
        try:
            os.rmdir(link)                 # remove so a ligacao, nunca o destino
        except OSError as exc:
            return False, (f"{link} e uma ligacao partida e nao consegui "
                           f"remove-la ({exc}). Apague-a a mao e repita.")

    try:
        os.makedirs(os.path.dirname(link), exist_ok=True)
    except OSError as exc:
        return False, f"nao consegui criar {os.path.dirname(link)}: {exc}"

    res = subprocess.run(['cmd', '/c', 'mklink', '/J', link, target],
                         capture_output=True, text=True)
    if res.returncode != 0 or not os.path.isdir(link):
        return False, (res.stderr or res.stdout or 'mklink falhou').strip()
    return True, f"{link}  ->  {target}"


def scan_for_stray_paths(model_dir: str) -> list[tuple[str, int, str]]:
    """Procura caminhos absolutos noutras tabelas (deviam usar o parametro)."""
    hits = []
    for dirpath, _dirs, files in os.walk(model_dir):
        for name in files:
            if not name.endswith(".tmdl") or name == "expressions.tmdl":
                continue
            full = os.path.join(dirpath, name)
            try:
                lines = open(full, encoding="utf-8").read().splitlines()
            except Exception:
                continue
            for i, line in enumerate(lines, 1):
                if ABS_PATH.search(line):
                    hits.append((os.path.relpath(full, model_dir), i, line.strip()))
    return hits


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    check_only = "--check" in argv

    tmdl = cfg.DASHBOARD_TMDL
    data_dir = cfg.DATA_POWERBI

    if not os.path.isfile(tmdl):
        print(f"ERRO: nao encontrei o ficheiro do modelo:\n  {tmdl}")
        return 1
    if not os.path.isdir(data_dir):
        print(f"AVISO: a pasta de dados ainda nao existe:\n  {data_dir}")

    # Preferimos o caminho estavel, igual em todos os PCs, para expressions.tmdl
    # deixar de mudar quando sincroniza. So caimos no caminho desta maquina se
    # a junction nao puder ser criada.
    target = data_dir
    if '--no-link' not in argv:
        ok, msg = ensure_junction(STABLE_LINK, data_dir)
        print(f"  Ligacao estavel   : {msg}")
        if ok:
            target = STABLE_LINK
        else:
            print("  -> a usar o caminho desta maquina. ATENCAO: assim o")
            print("     expressions.tmdl volta a ser especifico deste PC e")
            print("     sincroniza para os colegas.")

    text = open(tmdl, encoding="utf-8").read()
    actual = current_value(text)

    if actual is None:
        print("ERRO: nao encontrei o parametro DataFolder em expressions.tmdl.")
        return 1

    print(f"  DataFolder atual  : {actual}")
    print(f"  DataFolder correto: {target}")

    if os.path.normcase(actual) == os.path.normcase(target):
        print("  -> Ja esta correto para este PC. Nada a fazer.")
        rc = 0
    elif check_only:
        print("  -> ERRADO (modo --check: nao foi alterado).")
        rc = 1
    else:
        # Sem copia .bak_<data> ao lado: ficava na pasta PARTILHADA e
        # acumulava para sempre. O ficheiro esta em git — `git diff` mostra a
        # alteracao e `git checkout` desfaz.
        new_text = PATTERN.sub(lambda m: m.group(1) + target + m.group(3), text, count=1)
        with open(tmdl, "w", encoding="utf-8") as fh:
            fh.write(new_text)
        print("  -> CORRIGIDO.")
        print("     Abra o Power BI Desktop e clique em Atualizar (Refresh).")
        rc = 0

    model_dir = os.path.dirname(os.path.dirname(tmdl))
    stray = scan_for_stray_paths(model_dir)
    if stray:
        print()
        print("  AVISO: estas tabelas tem caminhos absolutos proprios em vez de")
        print("  usarem o parametro DataFolder - vao falhar noutros PCs:")
        for rel, line_no, line in stray:
            print(f"    {rel}:{line_no}  {line[:90]}")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
