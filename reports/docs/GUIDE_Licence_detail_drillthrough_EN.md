# Guide — Per-licence detail page (drill-through) — English

Goal: click a licence on the **Visão Geral** page and open a second page
**Ficha da Licença** already filtered to that `Nº TUA`, showing its parameters,
sampling/treatment obligations, and all the criteria to respect.

Power BI's UI here is in Portuguese, so button names are given in Portuguese with
the English meaning in brackets. All the data already exists in `Power BI Data\`
as `.xlsx` — you only need to re-link two tables and design the page.

---

## Part A — Bring `Conditions` + `Autocontrolo` into the model

They were removed when we pruned the model to the Visão Geral page. Add them back:

1. **Base (Home) → Obter dados (Get data) → Excel (pasta de trabalho / workbook)**.
2. Open `…\EPAL-project\Power BI Data\Conditions.xlsx` → in the **Navegador
   (Navigator)** tick sheet **Conditions** → **Transformar dados (Transform
   data)**, NOT "Carregar" (Load). Tick "Use first row as headers" if not already.
3. In Power Query, set the numeric types (otherwise they error or keep a `.0`):
   - `VLE mín`, `VLE máx`, `Carga máx. admissível (kg/dia)` → **Número decimal
     (Decimal number)**. If it errors, use **Usando definições regionais (Using
     locale)** → Decimal, locale **English (US)** (the decimal points are English).
   - The 7 criteria flag columns (`≤ 100% VLE (dobro)`, `≤ 150% VLE`,
     `≤ uma ordem de grandeza do VLE`, `média mensal ≤ VLE`, `média anual ≤ VLE`,
     `Quadro III DL 152/97 (borlas)`, `Gama de valores (intervalo)`) → leave as
     **Whole number (Número inteiro)** (they hold 0/1) or Text — either is fine.
4. Rename the query (right pane, **Nome / Name**) to **Conditions**.
5. Repeat steps 1–4 for `Autocontrolo.xlsx` (sheet **Autocontrolo**) → name
   **Autocontrolo**. Here set `Nº análises requeridas` → **Whole number**.
6. **Base (Home) → Fechar e aplicar (Close & Apply)**.

> Tip: to make these queries use the same folder parameter as the others, in the
> Advanced Editor of each, replace the fixed path with
> `Excel.Workbook(File.Contents(DataFolder & "\Conditions.xlsx"), null, true)`
> then `Source{[Item="Conditions",Kind="Sheet"]}[Data]` — exactly like the
> `Licenses` query already does.

---

## Part B — Relationships (star schema)

1. Go to **Model** view (linked-tables icon on the left).
2. Drag `Conditions[Nº TUA]` onto `Licenses[Nº TUA]`.
   - It should create a **Many-to-one (\*:1)**, **single** cross-filter direction,
     Conditions → Licenses. If Power BI proposes something else, double-click the
     relationship and confirm: *Cardinality = Many to one*, *Cross-filter
     direction = Single*.
3. Drag `Autocontrolo[Nº TUA]` onto `Licenses[Nº TUA]` — same thing.

Result: `Licenses` at the centre, with `Dim_Estado`, `Conditions` and
`Autocontrolo` around it. `Licenses` filters the others, never the reverse.

---

## Part C — Make a new page a drill-through target

1. At the bottom, click **+** for a new page. Double-click the name →
   **Ficha da Licença** (Licence sheet).
2. With the empty page selected, in the **Visualizations** pane find the
   **Drill-through (Fazer drill-through)** box. Drag `Licenses[Nº TUA]` into
   **"Add drill-through fields here"**.
   - This is the step that turns the page into a drill-through target: it opens
     filtered by the `Nº TUA` you came from.
3. Leave **"Keep all filters (Manter todos os filtros)"** on (optional).
4. A **back-arrow button** appears automatically — the "return" to Visão Geral.
5. Recommended: right-click the **Ficha da Licença** page tab → **Ocultar página
   (Hide page)**, so it only opens via drill-through.

---

## Part D — The four visuals

### 1. Header — licence identity and status
Insert a **Multi-row card (Cartão de várias linhas)** or several **Cards
(Cartões)**, and drag from `Licenses`:
`Estabelecimento`, `Nº TUA`, `Estado da Licença`, `Data de Validade`,
`Nível de tratamento`, `Esquema de tratamento`.
> Because the page is filtered to one `Nº TUA`, each card shows that licence's
> value. `Nível/Esquema de tratamento` covers the "treatments" part.

### 2. Table — Parameters and VLE
Insert a **Table (Tabela)** and, from `Conditions`, add in this order:
`Parâmetro`, `VLE`, `VLE mín`, `VLE máx`, `Carga máx. admissível (kg/dia)`,
`Frequência de amostragem`, `Tipo de amostragem`, `Legislação aplicável`.
> The limits and frequency the licence imposes on each parameter.

### 3. Table — Self-monitoring obligations (sampling)
Insert another **Table** and, from `Autocontrolo`, add:
`Local de amostragem`, `Parâmetro`, `Frequência de amostragem`,
`Tipo de amostragem`, `Nº análises requeridas`.
> The monitoring plan: where, what, how often, and how many analyses per year.

### 4. Matrix — Criteria to respect per parameter
Insert a **Matrix (Matriz)**:
- **Rows (Linhas)** → `Conditions[Parâmetro]`.
- **Values (Valores)** → the 7 criteria columns from `Conditions`:
  `média mensal ≤ VLE`, `média anual ≤ VLE`, `≤ 100% VLE (dobro)`,
  `≤ 150% VLE`, `≤ uma ordem de grandeza do VLE`, `Gama de valores (intervalo)`,
  `Quadro III DL 152/97 (borlas)`.
- Each value is 0/1. For each, click the field arrow → **Não resumir (Don't
  summarize)** (or *Maximum*) so it doesn't sum.

To show a ✓ instead of 1:
- Select the matrix → **Formatar (Format)** → **Elementos da célula (Cell
  elements)** → turn on **Ícones (Icons)** → **Formatação condicional
  (Conditional formatting)** per column → rule: *if value = 1 → green ✓; else
  blank*.
- Simpler alternative: instead of the flag matrix, use a table with `Parâmetro`
  + `Avaliação da Conformidade Legal` (the full legal text) + `Estado do critério`.

---

## Part E — Make the click work from Visão Geral

With the drill-through field set (Part C), it already works:

1. On **Visão Geral**, right-click a licence row (or a map bubble) → **Fazer
   drill-through (Drill through) → Ficha da Licença**.
2. The page opens filtered to that licence. The back arrow returns.

Optional — a permanent button instead of the right-click menu:
1. **Inserir (Insert) → Botões (Buttons) → Em branco (Blank)**. Text: "View
   licence detail".
2. Select the button → **Formatar botão (Format button) → Ação (Action) → Type =
   Drill-through → Destination = Ficha da Licença**.
3. The button is active only when a licence is selected in the table.

---

## Part F — Final check

- **Base (Home) → Atualizar (Refresh)** to import Conditions and Autocontrolo.
- Test 2–3 licences, including a **LURH**, a **TUA**, and a **caducada (expired)** one.
- Confirm the criteria matrix changes parameters between licences.
- An empty Autocontrolo table is normal for some licences — not all have a
  self-monitoring plan, and **LURH licences have no Autocontrolo rows in the
  combined feed at all** (only TUA does).

---

## Project-specific notes

- **Column sources**: parameters, VLE and criteria come from `Conditions.xlsx`
  (1207 rows); sampling comes from `Autocontrolo.xlsx` (953 rows, TUA only);
  treatment is on `Licenses`.
- **TUA and LURH together**: both origins are already in these files (`TUA/LURH`
  column in Licenses), so the page works the same for both — except Autocontrolo,
  which is TUA-only today.
- **Don't run `build_geral_page.py`** afterwards: the Visão Geral page is now
  hand-maintained; the script would recreate its visuals.
- **Want me to build this for you** in the project files instead of by hand? Say
  the word — I can generate the page and relationships and validate end-to-end.
