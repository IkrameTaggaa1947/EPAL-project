# -*- coding: utf-8 -*-
"""Limpar o que sobra de execucoes anteriores: um '<master>.PENDING.xlsx'
por integrar, e notas de revisao cujo PDF ja nao esta la.

PORQUE EXISTE ESTE FICHEIRO
---------------------------
Quando o master esta aberto no Excel, o pipeline nao consegue escrever nele e
desvia os dados para '<master>.PENDING.xlsx' - assim nada se perde. Faltava a
outra metade: ninguem voltava a ler esse ficheiro. Os dados ficavam guardados
num sitio de que o sistema nao tinha memoria, e a unica recuperacao possivel
era alguem reparar no ficheiro E ir buscar o PDF a data/review para o repor na
pasta de entrada.

reclaim() corre no arranque de cada pipeline:

  - nao ha PENDING            -> nao faz nada;
  - ha PENDING e o master ja
    esta livre                -> funde as linhas no master e apaga o PENDING;
  - ha PENDING mas o master
    continua bloqueado        -> NAO toca em nada e devolve a explicacao,
                                 para o pipeline a poder registar no log.

A fusao usa a mesma regra do upsert normal: as linhas antigas de cada licenca
presente no PENDING sao substituidas, nunca duplicadas.
"""
from __future__ import annotations

import os

import pandas as pd


def path_for(master_path: str) -> str:
    """O ficheiro PENDING que corresponde a este master."""
    return os.path.splitext(master_path)[0] + '.PENDING.xlsx'


def _chaves(folhas: dict, key: str) -> set[str]:
    """Todas as licencas presentes num conjunto de folhas {nome: DataFrame}."""
    return {str(v).strip() for df in folhas.values() if key in df.columns
            for v in df[key].dropna().astype(str) if str(v).strip()}


def reclaim(master_path: str, key: str, style=None) -> tuple[int, str]:
    """Funde '<master>.PENDING.xlsx' de volta no master.

    `key` e a coluna que identifica a licenca ('Nº TUA' / 'Nº Licença').
    `style` e opcional: recebe (worksheet, dataframe) e formata a folha - o
    extrator TUA passa a sua, para o master nao perder a formatacao.

    Devolve (linhas_integradas, mensagem). A mensagem e '' quando nao havia
    nada a fazer; caso contrario e uma frase para o log/operador.
    """
    pend = path_for(master_path)
    if not os.path.isfile(pend):
        return 0, ''
    nome = os.path.basename(pend)

    try:
        novo = pd.read_excel(pend, sheet_name=None)
    except Exception as exc:                       # corrompido / a sincronizar
        return 0, f"{nome} existe mas nao pode ser lido ({exc}) - deixado como esta."

    chaves = _chaves(novo, key)
    if not chaves:
        os.remove(pend)
        return 0, f"{nome} nao tinha linhas com {key} - removido."

    try:
        antigo = pd.read_excel(master_path, sheet_name=None) if os.path.isfile(master_path) else {}
    except Exception:
        return 0, (f"{nome} tem dados POR INTEGRAR, mas o master continua bloqueado. "
                   f"Feche o Excel e volte a correr.")

    # So se integram as licencas que FALTAM no master. Um PENDING pode ser mais
    # VELHO do que o master: se a mesma licenca foi reprocessada com exito
    # depois do desvio, a versao boa e a do master. Nunca sobrepor o master com
    # dados de um ficheiro cuja idade nao conhecemos.
    em_falta = chaves - _chaves(antigo, key)
    if not em_falta:
        os.remove(pend)
        return 0, (f"{nome} ja estava todo no master ({len(chaves)} licenca(s)) - "
                   f"removido sem alterar nada.")

    fundido = {}
    linhas = 0
    for name in list(dict.fromkeys(list(antigo) + list(novo))):
        velho = antigo.get(name, pd.DataFrame())
        recente = novo.get(name, pd.DataFrame())
        if not recente.empty and key in recente.columns:
            recente = recente[recente[key].astype(str).isin(em_falta)]
        elif not recente.empty:
            recente = recente.iloc[0:0]            # folha sem chave: nada a integrar
        fundido[name] = (pd.concat([velho, recente], ignore_index=True)
                         if len(velho) or len(recente) else velho)
        linhas += len(recente)

    try:
        with pd.ExcelWriter(master_path, engine='openpyxl') as xl:
            for name, df in fundido.items():
                folha = name[:31]
                out = df if not df.empty else pd.DataFrame({'(sem dados)': []})
                out.to_excel(xl, sheet_name=folha, index=False)
                if style is not None:
                    style(xl.sheets[folha], out)
    except PermissionError:
        return 0, (f"{nome} tem dados POR INTEGRAR, mas o master esta aberto no Excel. "
                   f"Feche-o e volte a correr.")

    os.remove(pend)
    return linhas, (f"{nome} integrado no master ({linhas} linhas, "
                    f"{len(chaves)} licenca(s)) e removido.")


def sweep_orphan_notes(review_dir: str) -> tuple[int, str]:
    """Apaga os '<nome>_REVIEW.txt' cujo PDF ja nao esta em review_dir.

    A propria nota diz ao operador para mover o PDF para o arquivo quando os
    dados estiverem certos - mas nada apagava a nota, por isso a pasta enchia-se
    de avisos sobre ficheiros que ja la nao estao (eram 70 para 2 PDFs). Uma
    nota sem o seu PDF nao tem nada que se veja: e ruido numa pasta que devia
    ser uma lista de trabalho.

    Devolve (quantas_apagadas, mensagem).
    """
    if not os.path.isdir(review_dir):
        return 0, ''
    apagadas = []
    for name in sorted(os.listdir(review_dir)):
        if not name.endswith('_REVIEW.txt'):
            continue
        base = name[:-len('_REVIEW.txt')]
        if any(os.path.isfile(os.path.join(review_dir, base + ext))
               for ext in ('.pdf', '.PDF')):
            continue                              # o PDF ainda ca esta
        try:
            os.remove(os.path.join(review_dir, name))
            apagadas.append(name)
        except OSError:
            pass                                  # bloqueada/a sincronizar: fica
    if not apagadas:
        return 0, ''
    return len(apagadas), (f"{len(apagadas)} nota(s) de revisao orfa(s) removida(s) "
                           f"(o PDF correspondente ja nao esta em review/).")


# ------------------------------------------------------------------ self-check
def _demo():
    """As duas regras que interessam:
       1. uma licenca que FALTA no master e integrada;
       2. uma licenca que JA esta no master nao e sobreposta - o PENDING pode
          ser mais velho do que o master (foi o caso real que motivou isto)."""
    import tempfile
    key = 'Nº TUA'

    def escrever(path, linhas):
        with pd.ExcelWriter(path, engine='openpyxl') as xl:
            pd.DataFrame(linhas).to_excel(xl, sheet_name='Licenses', index=False)

    def ler(path):
        got = pd.read_excel(path, sheet_name='Licenses')
        return dict(zip(got[key].astype(str), got['v'].astype(str)))

    with tempfile.TemporaryDirectory() as tmp:
        master = os.path.join(tmp, 'master.xlsx')

        # 1. licenca em falta -> integrada, as outras ficam intactas
        escrever(master, [{key: 'B', 'v': 'intacto'}])
        escrever(path_for(master), [{key: 'A', 'v': 'recuperado'}])
        linhas, msg = reclaim(master, key)
        assert linhas == 1, (linhas, msg)
        assert not os.path.isfile(path_for(master)), 'o PENDING devia ter sido removido'
        assert ler(master) == {'A': 'recuperado', 'B': 'intacto'}, ler(master)

        # 2. licenca ja no master -> PENDING descartado, master NAO alterado
        escrever(path_for(master), [{key: 'A', 'v': 'versao velha'}])
        linhas, msg = reclaim(master, key)
        assert linhas == 0, (linhas, msg)
        assert not os.path.isfile(path_for(master)), 'o PENDING devia ter sido removido'
        assert ler(master) == {'A': 'recuperado', 'B': 'intacto'}, ler(master)

        # 3. sem PENDING nao faz nada e nao se queixa
        assert reclaim(master, key) == (0, '')

    # 4. notas orfas: so desaparece a que perdeu o PDF
    with tempfile.TemporaryDirectory() as rev:
        def toca(nome):
            with open(os.path.join(rev, nome), 'w', encoding='utf-8') as fh:
                fh.write('x')
        toca('tem_pdf.pdf'); toca('tem_pdf_REVIEW.txt')
        toca('sem_pdf_REVIEW.txt')
        n, msg = sweep_orphan_notes(rev)
        assert n == 1, (n, msg)
        restantes = set(os.listdir(rev))
        assert restantes == {'tem_pdf.pdf', 'tem_pdf_REVIEW.txt'}, restantes
        assert sweep_orphan_notes(rev) == (0, '')      # idempotente
    print('pending self-check OK')


if __name__ == '__main__':
    _demo()
