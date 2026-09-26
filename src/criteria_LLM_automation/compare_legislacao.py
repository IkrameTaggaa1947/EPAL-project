# -*- coding: utf-8 -*-
"""Comparar a legislação/critérios de uma licença ANTIGA com a NOVA (mesma ETAR).

Objetivo (EPAL): quando chega uma licença nova para uma ETAR que já tinha licença,
verificar se a COMBINAÇÃO DE CRITÉRIOS a cumprir (critérios de conformidade por
parâmetro) muda ou não. A parte difícil é semântica — a lei citada é reescrita de
licença para licença mesmo quando quer dizer o mesmo — e é aí que entra o AMALIA
(LLM português). Ver amalia_client.py para o acesso ao modelo.

O que este script faz
----------------------
1. Empareha a licença antiga e a nova pela mesma ETAR (via CODMAXIMO em
   data/powerbi/Licenses.xlsx). Por omissão usa a ETAR Ade (a única no conjunto
   atual com licença antiga LURH + nova TUA).
2. Para cada parâmetro (CBO5, CQO, SST, pH, ...), recolhe de data/powerbi:
      VLE, Legislação aplicável, Avaliação da Conformidade Legal   (Conditions.xlsx)
      combinação de critérios = IdCriterio -> vetor de flags        (Licencas_Criterios + Catalogo_Criterios)
3. Diferença DETERMINÍSTICA: parâmetro removido/novo, VLE mudou, combinação de
   critérios mudou (flags adicionadas/removidas) — quando ambas as combinações são
   conhecidas nas tabelas.
4. Camada SEMÂNTICA (AMALIA): quando o texto da legislação difere, pede ao modelo
   para dizer se são equivalentes e se a combinação de critérios muda, com
   justificação jurídica. Sem credenciais, escreve o Excel na mesma e deixa estas
   colunas a "PENDENTE" (o prompt fica guardado na folha Prompts_AMALIA).
5. Escreve um Excel: Resumo + Comparação + Prompts_AMALIA + Leia-me.

Uso:
    python compare_legislacao.py                         # ETAR Ade (por omissão)
    python compare_legislacao.py --cod 3-SBA-ADEE-...    # outra ETAR por CODMAXIMO
    python compare_legislacao.py --old <NºLic> --new <NºTUA>
    python compare_legislacao.py --out caminho.xlsx
"""
from __future__ import annotations
import argparse
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.dirname(HERE)                                   # …/EPAL-project/src
sys.path.insert(0, HERE)
sys.path.insert(0, SRC)

import epal_config                                            # noqa: E402
# Carrega as definicoes do AMALIA a partir de %LOCALAPPDATA%\EPAL\epal.local.ini
# (proprias de cada maquina) para as variaveis de ambiente que o cliente le.
epal_config.apply_amalia_env()

import amalia_client                                          # noqa: E402
import textos_distintos                                       # noqa: E402

ROOT = epal_config.ROOT                                       # …/EPAL-project
PBI = epal_config.DATA_POWERBI

DEFAULT_COD = "3-SBA-ADEE-RSADE-ETRADE"                       # ETAR Ade

# Ordem e rótulos legíveis das 9 flags de critério (como em Catalogo_Criterios).
FLAG_COLS = [
    "VLE (flag)", "% mín remoção (flag)", "Gama de valores (intervalo)",
    "≤ 100% VLE (dobro)", "≤ 150% VLE", "≤ uma ordem de grandeza do VLE",
    "média mensal ≤ VLE", "média anual ≤ VLE", "Quadro III DL 152/97 (borlas)",
]
FLAG_LABEL = {
    "VLE (flag)": "VLE",
    "% mín remoção (flag)": "% mín. remoção",
    "Gama de valores (intervalo)": "Gama de valores",
    "≤ 100% VLE (dobro)": "≤100% VLE (dobro)",
    "≤ 150% VLE": "≤150% VLE",
    "≤ uma ordem de grandeza do VLE": "≤1 ordem de grandeza",
    "média mensal ≤ VLE": "média mensal ≤ VLE",
    "média anual ≤ VLE": "média anual ≤ VLE",
    "Quadro III DL 152/97 (borlas)": "Quadro III DL 152/97",
}


# --------------------------------------------------------------------------- #
#  Leitura das folhas
# --------------------------------------------------------------------------- #
def _load(fname):
    wb = openpyxl.load_workbook(os.path.join(PBI, fname), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    return list(rows[0]), rows[1:]


def _col(header, name):
    for j, x in enumerate(header):
        if x == name:
            return j
    return None


def _canon_param(texto):
    """Nome canónico curto a partir do rótulo longo do documento."""
    t = (texto or "").lower()
    table = [
        ("cbo", "CBO5"), ("bioquímica", "CBO5"), ("bioquimica", "CBO5"),
        ("cqo", "CQO"), ("química", "CQO"), ("quimica", "CQO"),
        ("sólidos suspensos", "SST"), ("solidos suspensos", "SST"), ("sst", "SST"),
        ("ph", "pH"),
        ("azoto amoniacal", "N-NH4"), ("amoniacal", "N-NH4"),
        ("azoto total", "N-total"), ("azoto", "N-total"),
        ("fósforo", "P-total"), ("fosforo", "P-total"),
        ("coliformes", "Coliformes"), ("óleos", "Óleos e gorduras"),
    ]
    for key, canon in table:
        if key in t:
            return canon
    return (texto or "").strip()


def _norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


# --------------------------------------------------------------------------- #
#  Modelo de dados
# --------------------------------------------------------------------------- #
def licence_index():
    h, r = _load("Licenses.xlsx")
    return h, r, {
        "id": _col(h, "Nº TUA"), "tipo": _col(h, "TUA/LURH"),
        "est": _col(h, "Estabelecimento"), "cod": _col(h, "CODMAXIMO"),
        "vig": _col(h, "Data de Entrada em Vigor"), "val": _col(h, "Data de Validade"),
    }


def _to_date(s):
    m = re.match(r"(\d{2})-(\d{2})-(\d{4})", str(s or ""))
    return (int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else (0, 0, 0)


def pick_pair(cod=None, old=None, new=None):
    """Devolve (old_lic, new_lic) — cada um dict com id/tipo/est/vig/val."""
    h, rows, ix = licence_index()
    def rec(r):
        return {"id": str(r[ix["id"]]), "tipo": r[ix["tipo"]], "est": r[ix["est"]],
                "cod": r[ix["cod"]], "vig": r[ix["vig"]], "val": r[ix["val"]]}
    if old and new:
        byid = {str(r[ix["id"]]): rec(r) for r in rows}
        return byid[old], byid[new]
    cod = cod or DEFAULT_COD
    grp = [rec(r) for r in rows if str(r[ix["cod"]]) == str(cod)]
    if len(grp) < 2:
        raise SystemExit(f"CODMAXIMO {cod}: encontrei {len(grp)} licença(s) — preciso de "
                         f"uma antiga e uma nova. Use --old/--new.")
    grp.sort(key=lambda d: _to_date(d["vig"]))
    return grp[0], grp[-1]                          # mais antiga por vigor, mais recente


def conditions_by_licence(lic_id):
    """{param_canonico: {'raw','vle','leg','aval'}} para uma licença."""
    h, rows = _load("Conditions.xlsx")
    ck, cp = _col(h, "Nº TUA"), _col(h, "Parâmetro") or _col(h, "Parametro")
    cv, cl, ca = _col(h, "VLE"), _col(h, "Legislação aplicável"), _col(h, "Avaliação da Conformidade Legal")
    out = {}
    for r in rows:
        if str(r[ck]) != str(lic_id):
            continue
        canon = _canon_param(r[cp])
        out[canon] = {"raw": _norm(r[cp]), "vle": _norm(r[cv]),
                      "leg": _norm(r[cl]), "aval": _norm(r[ca])}
    return out


def _catalogo():
    h, rows = _load("Catalogo_Criterios.xlsx")
    idc = _col(h, "IdCriterio")
    fi = [_col(h, c) for c in FLAG_COLS]
    cat = {}
    for r in rows:
        active = [FLAG_COLS[k] for k, j in enumerate(fi)
                  if j is not None and str(r[j]) not in ("0", "None", "", "0.0")]
        cat[r[idc]] = active
    return cat


def criteria_by_licence(lic_id, cat):
    """{param_canonico: {'idc','crit':set(labels)}} — só onde a licença está na
    tabela de factos Licencas_Criterios (as TUA estão; muitas LURH antigas não)."""
    h, rows = _load("Licencas_Criterios.xlsx")
    li, lp, lc = _col(h, "IdLicenca"), _col(h, "Parametro"), _col(h, "IdCriterio")
    out = {}
    for r in rows:
        if str(r[li]) != str(lic_id):
            continue
        canon = _canon_param(r[lp])
        active = cat.get(r[lc], [])
        out[canon] = {"idc": r[lc], "crit": set(FLAG_LABEL[c] for c in active)}
    return out


# --------------------------------------------------------------------------- #
#  Prompt para o AMALIA
# --------------------------------------------------------------------------- #
SYSTEM = ("És um jurista especialista em direito do ambiente português, em concreto "
          "nas normas de descarga de águas residuais (Lei n.º 58/2005 - Lei da Água, "
          "Decreto-Lei n.º 236/98, Decreto-Lei n.º 152/97 e legislação conexa). "
          "Respondes com rigor, apenas com base nos textos fornecidos, em português "
          "de Portugal e no formato pedido.")

VOCAB = ("VLE; % mín. remoção; Gama de valores; ≤100% VLE (dobro); ≤150% VLE; "
         "≤1 ordem de grandeza do VLE; média mensal ≤ VLE; média anual ≤ VLE; "
         "Quadro III do DL 152/97 (contagem estatística de amostras não conformes / 'borlas')")


def build_prompt(etar, param, old, new):
    return f"""Compara, para a mesma ETAR ({etar}) e o mesmo parâmetro ({param}), a base legal e o método de avaliação da conformidade de uma licença ANTIGA e de uma licença NOVA.

=== LICENÇA ANTIGA ===
VLE: {old.get('vle','(n/d)')}
Legislação aplicável: {old.get('leg','(n/d)')}
Avaliação da conformidade: {old.get('aval','(n/d)')}

=== LICENÇA NOVA ===
VLE: {new.get('vle','(n/d)')}
Legislação aplicável: {new.get('leg','(n/d)')}
Avaliação da conformidade: {new.get('aval','(n/d)')}

Vocabulário de critérios a usar (escolhe apenas destes):
{VOCAB}

Responde EXATAMENTE neste formato, sem texto adicional:
EQUIVALENCIA_SEMANTICA: <Sim|Não|Parcial>
CRITERIOS_ANTIGOS: <critérios do vocabulário aplicáveis à licença antiga, separados por ';'>
CRITERIOS_NOVOS: <critérios do vocabulário aplicáveis à licença nova, separados por ';'>
COMBINACAO_MUDA: <Sim|Não>
JUSTIFICACAO: <UMA frase curta>"""


def parse_amalia(text):
    out = {"EQUIVALENCIA_SEMANTICA": "", "CRITERIOS_ANTIGOS": "",
           "CRITERIOS_NOVOS": "", "COMBINACAO_MUDA": "", "JUSTIFICACAO": ""}
    for line in (text or "").splitlines():
        for k in out:
            if line.strip().upper().startswith(k):
                out[k] = line.split(":", 1)[1].strip() if ":" in line else ""
    return out


# --------------------------------------------------------------------------- #
#  Comparação
# --------------------------------------------------------------------------- #
def compare(old_lic, new_lic):
    cat = _catalogo()
    o_cond, n_cond = conditions_by_licence(old_lic["id"]), conditions_by_licence(new_lic["id"])
    o_crit, n_crit = criteria_by_licence(old_lic["id"], cat), criteria_by_licence(new_lic["id"], cat)
    params = sorted(set(o_cond) | set(n_cond))
    etar = new_lic["est"] or old_lic["est"]
    rows = []
    for p in params:
        o, n = o_cond.get(p), n_cond.get(p)
        oc = o_crit.get(p, {}).get("crit") if o else None
        nc = n_crit.get(p, {}).get("crit") if n else None

        if o and not n:
            estado = "PARÂMETRO REMOVIDO"
        elif n and not o:
            estado = "PARÂMETRO NOVO"
        elif oc is not None and nc is not None:
            estado = "Sem alteração (combinação idêntica)" if oc == nc else "COMBINAÇÃO ALTERA"
        else:
            leg_diff = (_norm(o["leg"]) != _norm(n["leg"])) if (o and n) else True
            estado = ("COMBINAÇÃO A CONFIRMAR (texto legal difere)" if leg_diff
                      else "Provável sem alteração (texto legal idêntico)")

        det_delta = ""
        if oc is not None and nc is not None and oc != nc:
            add = sorted(nc - oc); rem = sorted(oc - nc)
            det_delta = " | ".join(x for x in (("+ " + "; ".join(add)) if add else "",
                                               ("- " + "; ".join(rem)) if rem else "") if x)

        prompt = build_prompt(etar, p, o or {}, n or {}) if (o and n) else ""
        rows.append({
            "param": p,
            "vle_o": (o or {}).get("vle", ""), "vle_n": (n or {}).get("vle", ""),
            "leg_o": (o or {}).get("leg", ""), "leg_n": (n or {}).get("leg", ""),
            "aval_o": (o or {}).get("aval", ""), "aval_n": (n or {}).get("aval", ""),
            "crit_o": "; ".join(sorted(oc)) if oc else ("(não nas tabelas — inferir pelo AMALIA)" if o else "—"),
            "crit_n": "; ".join(sorted(nc)) if nc else ("(não nas tabelas — inferir pelo AMALIA)" if n else "—"),
            "estado": estado, "det_delta": det_delta, "prompt": prompt,
        })
    return etar, rows


def run_amalia(rows):
    """Preenche colunas do AMALIA in-place. Devolve (usou_modelo, backend)."""
    if not amalia_client.is_configured():
        for r in rows:
            r.update({"am_equiv": "PENDENTE", "am_muda": "PENDENTE",
                      "am_crit_o": "", "am_crit_n": "", "am_just": ""})
        return False, amalia_client.backend_name()
    # UMA pergunta por PAR DE TEXTOS, nao uma por linha. A pergunta semantica
    # nao depende da ETAR nem do parametro -- depende so dos dois textos legais
    # -- e os textos repetem-se muito (1810 linhas -> 129 pares distintos no
    # conjunto actual). Com o modelo local a 0,8 tokens/s medidos, isto e a
    # diferenca entre ~314 horas e ~22. As respostas ficam em cache no disco:
    # uma execucao interrompida retoma onde ia e a segunda passagem nao paga nada.
    CAMPOS = ("am_equiv", "am_muda", "am_crit_o", "am_crit_n", "am_just")
    cache = textos_distintos.Cache()

    for r in rows:                              # sem prompt = nada a perguntar
        if not r["prompt"]:
            r.update({"am_equiv": "n/a", "am_muda": "n/a",
                      "am_crit_o": "", "am_crit_n": "", "am_just": ""})

    por_perguntar = [r for r in rows if r["prompt"]]
    grupos = textos_distintos.pares_distintos(por_perguntar)
    novos = [k for k in grupos if cache.get(k) is None]
    print('  AMALIA: %d linha(s) -> %d par(es) distinto(s) '
          '(%d em cache, %d por perguntar)'
          % (len(por_perguntar), len(grupos), len(grupos) - len(novos), len(novos)))

    for i, k in enumerate(novos, 1):
        exemplo = grupos[k][0]
        print('     [%d/%d] %s ...' % (i, len(novos), str(exemplo.get('param', '?'))[:30]),
              flush=True)
        try:
            ans = parse_amalia(amalia_client.ask(exemplo["prompt"], system=SYSTEM))
            cache.put(k, {"am_equiv": ans["EQUIVALENCIA_SEMANTICA"],
                          "am_muda": ans["COMBINACAO_MUDA"],
                          "am_crit_o": ans["CRITERIOS_ANTIGOS"],
                          "am_crit_n": ans["CRITERIOS_NOVOS"],
                          "am_just": ans["JUSTIFICACAO"]})
        except Exception as exc:                # nunca aborta o Excel por erro do modelo
            # NAO vai para a cache: um erro de rede nao pode ficar gravado como
            # se fosse a resposta do modelo, senao nunca mais se volta a tentar.
            print('        ERRO: %s' % str(exc)[:70])
            for r in por_perguntar:
                if textos_distintos.chave(r.get("leg_o"), r.get("aval_o"),
                                          r.get("leg_n"), r.get("aval_n")) == k:
                    r.update({"am_equiv": "ERRO: %s" % exc, "am_muda": "",
                              "am_crit_o": "", "am_crit_n": "", "am_just": ""})

    textos_distintos.aplicar(por_perguntar, cache, CAMPOS)
    return True, amalia_client.backend_name()


# --------------------------------------------------------------------------- #
#  Excel
# --------------------------------------------------------------------------- #
BOLD = Font(bold=True)
WHITE = Font(bold=True, color="FFFFFF")
HDR = PatternFill("solid", fgColor="0D6ABF")
CH_FILL = PatternFill("solid", fgColor="FFE0B2")     # laranja claro = alterou
OK_FILL = PatternFill("solid", fgColor="E8F5E9")     # verde claro = sem alteração
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(*[Side(style="thin", color="D9D9D9")] * 4)


def _style_header(ws, ncols, row=1):
    for j in range(1, ncols + 1):
        c = ws.cell(row=row, column=j)
        c.font = WHITE; c.fill = HDR; c.alignment = WRAP; c.border = THIN


def write_excel(path, etar, old_lic, new_lic, rows, used_model, backend):
    wb = openpyxl.Workbook()

    # ---- Resumo -----------------------------------------------------------
    ws = wb.active; ws.title = "Resumo"
    muda = any(r["estado"].startswith(("PARÂMETRO", "COMBINAÇÃO ALTERA", "COMBINAÇÃO A CONFIRMAR"))
               for r in rows)
    veredito = ("OS CRITÉRIOS A CUMPRIR MUDAM (confirmar com o AMALIA)" if muda
                else "Sem alterações aparentes")
    info = [
        ("Comparação de legislação/critérios — licença antiga vs nova", ""),
        ("", ""),
        ("ETAR", etar),
        ("CODMAXIMO", old_lic.get("cod") or new_lic.get("cod")),
        ("Licença ANTIGA", f'{old_lic["id"]} ({old_lic["tipo"]})  vigor {old_lic["vig"]} → validade {old_lic["val"]}'),
        ("Licença NOVA", f'{new_lic["id"]} ({new_lic["tipo"]})  vigor {new_lic["vig"]} → validade {new_lic["val"]}'),
        ("Nº de parâmetros (antiga / nova)",
         f'{sum(1 for r in rows if r["crit_o"]!="—")} / {sum(1 for r in rows if r["crit_n"]!="—")}'),
        ("Veredito determinístico", veredito),
        ("Análise semântica AMALIA", "aplicada" if used_model else f"PENDENTE — back-end {backend} (definir credenciais)"),
        ("", ""),
        ("Nota", "A camada determinística compara o que já está nas tabelas de critérios. "
                 "A coluna AMALIA resolve os casos em que o texto legal muda de forma mas "
                 "não (necessariamente) de conteúdo. Ver folha 'Comparação'."),
    ]
    for i, (k, v) in enumerate(info, 1):
        ws.cell(row=i, column=1, value=k).font = BOLD
        ws.cell(row=i, column=2, value=v).alignment = WRAP
    ws.cell(row=8, column=2).font = Font(bold=True, color="C0392B" if muda else "1E8449")
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 95

    # ---- Comparação -------------------------------------------------------
    ws = wb.create_sheet("Comparação")
    cols = [
        ("Parâmetro", 16), ("VLE antiga", 12), ("VLE nova", 12),
        ("Combinação de critérios — ANTIGA", 34), ("Combinação de critérios — NOVA", 34),
        ("Estado (determinístico)", 30), ("Δ critérios (determinístico)", 28),
        ("AMALIA: equivalência", 16), ("AMALIA: combinação muda?", 18),
        ("AMALIA: critérios antigos", 26), ("AMALIA: critérios novos", 26),
        ("AMALIA: justificação", 50),
        ("Legislação — ANTIGA", 48), ("Legislação — NOVA", 48),
        ("Avaliação conformidade — ANTIGA", 50), ("Avaliação conformidade — NOVA", 50),
    ]
    for j, (name, w) in enumerate(cols, 1):
        ws.cell(row=1, column=j, value=name)
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    _style_header(ws, len(cols))
    for r in rows:
        vals = [r["param"], r["vle_o"], r["vle_n"], r["crit_o"], r["crit_n"],
                r["estado"], r["det_delta"], r.get("am_equiv", ""), r.get("am_muda", ""),
                r.get("am_crit_o", ""), r.get("am_crit_n", ""), r.get("am_just", ""),
                r["leg_o"], r["leg_n"], r["aval_o"], r["aval_n"]]
        ws.append(vals)
        rr = ws.max_row
        changed = not r["estado"].lower().startswith(("sem alteração", "provável sem"))
        for j in range(1, len(cols) + 1):
            c = ws.cell(row=rr, column=j); c.alignment = WRAP; c.border = THIN
            if j == 6:
                c.fill = CH_FILL if changed else OK_FILL
                c.font = BOLD
    ws.freeze_panes = "B2"; ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(cols))}1"

    # ---- Prompts_AMALIA ---------------------------------------------------
    ws = wb.create_sheet("Prompts_AMALIA")
    ws.append(["Parâmetro", "Prompt enviado ao AMALIA (system: jurista ambiental)"])
    _style_header(ws, 2)
    ws.column_dimensions["A"].width = 16; ws.column_dimensions["B"].width = 140
    for r in rows:
        if r["prompt"]:
            ws.append([r["param"], r["prompt"]])
            ws.cell(row=ws.max_row, column=2).alignment = WRAP

    # ---- Leia-me ----------------------------------------------------------
    ws = wb.create_sheet("Leia-me")
    doc = [
        "COMO LER ESTE FICHEIRO",
        "",
        "Resumo        — a ETAR, as duas licenças e o veredito global.",
        "Comparação    — uma linha por parâmetro. As colunas 'determinístico' vêm das tabelas",
        "                de critérios do projeto; as colunas 'AMALIA' vêm do LLM português.",
        "Prompts_AMALIA— o texto exato enviado ao modelo (transparência / uso manual).",
        "",
        "COMBINAÇÃO DE CRITÉRIOS = conjunto de critérios de conformidade que se aplicam a um",
        "parâmetro. Vocabulário: " + VOCAB,
        "",
        "LIGAR O AMALIA (preenche as colunas AMALIA ao correr de novo):",
        "  1. Abrir o ficheiro de definicoes DESTE PC:  %LOCALAPPDATA%\\EPAL\\epal.local.ini",
        "  2. Na seccao [amalia], preencher:",
        "       vLLM local:  backend = openai   endpoint = http://localhost:8000/v1   api_key = local",
        "       IAedu:       backend = iaedu    endpoint = <url do IAedu>             api_key = <chave>",
        "  3. Voltar a correr COMPARAR_CRITERIOS.bat",
        "  (O ficheiro e local a cada computador - nao e partilhado com os colegas.)",
        "",
        "Sem credenciais o ficheiro é gerado na mesma, com as colunas AMALIA a 'PENDENTE'.",
    ]
    for i, line in enumerate(doc, 1):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 120

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="Comparar legislação/critérios antiga vs nova (mesma ETAR).")
    ap.add_argument("--cod", help="CODMAXIMO da ETAR (por omissão: ETAR Ade).")
    ap.add_argument("--old", help="Nº da licença antiga (em vez de --cod).")
    ap.add_argument("--new", help="Nº da licença nova (em vez de --cod).")
    ap.add_argument("--out", help="Caminho do Excel de saída.")
    a = ap.parse_args(argv)

    old_lic, new_lic = pick_pair(cod=a.cod, old=a.old, new=a.new)
    etar, rows = compare(old_lic, new_lic)
    used_model, backend = run_amalia(rows)

    safe = re.sub(r"[^\w]+", "_", str(etar)).strip("_") or "ETAR"
    out = a.out or os.path.join(ROOT, "outputs", "comparacoes", f"Comparacao_Criterios_{safe}.xlsx")
    write_excel(out, etar, old_lic, new_lic, rows, used_model, backend)

    print(f"ETAR: {etar}")
    print(f"  Antiga: {old_lic['id']} ({old_lic['tipo']})   Nova: {new_lic['id']} ({new_lic['tipo']})")
    for r in rows:
        print(f"  - {r['param']:8} {r['estado']}"
              + (f"   [{r['det_delta']}]" if r['det_delta'] else ""))
    print(f"AMALIA: {'aplicado' if used_model else 'PENDENTE (' + backend + ')'}")
    print(f"Excel: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
