# -*- coding: utf-8 -*-
"""Garantir 'Id Versão' em TODAS as tabelas do feed que se ligam a uma licenca.

PORQUE
------
O Power BI exige que a coluna do lado UM de uma relacao seja UNICA. Desde que
uma renovacao passou a ser uma linha propria, o numero da licenca deixou de o
ser -- a ETAR Sao Miguel tem duas licencas com o numero TUA20250408001124 -- e
a atualizacao rebentava logo na tabela 'Licences', arrastando todas as outras:

    a coluna 'Nº Licence' contem um valor em duplicado 'TUA20250408001124'

'Id Versão' (nome + periodo) e unica, e e por ela que o modelo passa a
relacionar-se. As tabelas vem de sitios diferentes -- umas do master TUA, outras
do master LURH, outras de ficheiros feitos a mao -- e nem todas a trazem de
origem. Este passo fecha essa diferenca num sitio so, em vez de a remendar em
cada construtor de linhas.

O CASO AMBIGUO, DITO EM VOZ ALTA
--------------------------------
As tabelas cuja origem so conhece o NUMERO da licenca nao sabem, para o numero
repetido, a que versao pertence cada linha. Essas ficam ligadas a versao
marcada 'Atual'. E o mais fiel aos dados: a versao substituida nao tem
condicoes nem criterios acompanhados em separado.
"""
from __future__ import annotations

import os

import pandas as pd

COLUNA = 'Id Versão'

# (ficheiro, coluna que identifica a licenca)
TABELAS = [
    ('Conditions.xlsx', 'Nº TUA'),
    ('Autocontrolo.xlsx', 'Nº TUA'),
    ('Avaliacao.xlsx', 'Nº TUA'),
    ('Comunicacoes.xlsx', 'Nº TUA'),
    ('MeioRecetor.xlsx', 'Nº TUA'),
    ('Condicoes_Prioritarias.xlsx', 'Nº TUA'),
    ('Licencas_Criterios.xlsx', 'IdLicenca'),
    ('Pesquisa.xlsx', 'Nº TUA'),
]


def mapa_numero_para_versao(lic: pd.DataFrame) -> dict:
    """numero -> Id Versão. Se o numero se repetir, ganha a versao 'Atual'."""
    mapa = {}
    for _, r in lic.iterrows():
        num = str(r.get('Nº TUA', '')).strip()
        if not num:
            continue
        atual = str(r.get('Vigência da Licença', '')).strip() == 'Atual'
        if num not in mapa or atual:
            mapa[num] = str(r.get(COLUNA, '')).strip()
    return mapa


def completar(pasta_feed: str, log=print) -> int:
    """Preenche COLUNA onde faltar. Devolve quantas tabelas foram alteradas."""
    p_lic = os.path.join(pasta_feed, 'Licenses.xlsx')
    if not os.path.isfile(p_lic):
        return 0
    lic = pd.read_excel(p_lic, dtype=str).fillna('')
    if COLUNA not in lic.columns:
        log('  %s: Licenses.xlsx nao tem a coluna - nada a fazer' % COLUNA)
        return 0
    mapa = mapa_numero_para_versao(lic)

    alteradas = 0
    for nome, chave in TABELAS:
        p = os.path.join(pasta_feed, nome)
        if not os.path.isfile(p):
            continue
        # O NOME DA FOLHA tem de se manter: a particao do modelo procura-a pelo
        # nome (Source{[Item = "Conditions", Kind = "Sheet"]}). Um to_excel sem
        # sheet_name grava 'Sheet1' e o Power BI recusa carregar a tabela.
        folha = pd.ExcelFile(p).sheet_names[0]
        d = pd.read_excel(p, sheet_name=folha, dtype=str).fillna('')
        if chave not in d.columns:
            continue
        antes = d[COLUNA].astype(str).str.strip() if COLUNA in d.columns else None
        novo = d[chave].astype(str).str.strip().map(mapa).fillna('')
        # o que a tabela ja trazia manda; so se preenche o que estava vazio
        d[COLUNA] = novo if antes is None else antes.where(antes.ne(''), novo)
        with pd.ExcelWriter(p, engine='openpyxl') as xl:
            d.to_excel(xl, sheet_name=folha[:31], index=False)
        alteradas += 1
        vazias = int((d[COLUNA].astype(str).str.strip() == '').sum())
        log('  %-30s %5d linhas, %d sem licenca correspondente' % (nome, len(d), vazias))
    return alteradas


# ------------------------------------------------------------------ self-check
def _demo():
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        # duas licencas com o MESMO numero: a atual e a substituida
        pd.DataFrame([
            {'Nº TUA': 'T1', COLUNA: 'x|2026|2031', 'Vigência da Licença': 'Atual'},
            {'Nº TUA': 'T1', COLUNA: 'x|2023|2026', 'Vigência da Licença': 'Histórico'},
            {'Nº TUA': 'T2', COLUNA: 'y|2020|2030', 'Vigência da Licença': 'Atual'},
        ]).to_excel(os.path.join(tmp, 'Licenses.xlsx'), index=False)
        pd.DataFrame([{'Nº TUA': 'T1', 'v': 1}, {'Nº TUA': 'T2', 'v': 2},
                      {'Nº TUA': 'T9', 'v': 3}]).to_excel(
            os.path.join(tmp, 'Conditions.xlsx'), index=False)

        completar(tmp, log=lambda *a: None)
        d = pd.read_excel(os.path.join(tmp, 'Conditions.xlsx'), dtype=str).fillna('')
        got = dict(zip(d['Nº TUA'], d[COLUNA]))
        assert got['T1'] == 'x|2026|2031', got      # numero repetido -> a ATUAL
        assert got['T2'] == 'y|2020|2030', got
        assert got['T9'] == '', got                 # licenca que nao existe -> vazio
    print('completar_id_versao self-check OK')


if __name__ == '__main__':
    _demo()
