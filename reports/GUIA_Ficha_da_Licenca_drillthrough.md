# Guia — Página de detalhe por licença (drill-through)

Objetivo: clicar numa licença na página **Visão Geral** e abrir uma segunda
página **Ficha da Licença** já filtrada por esse `Nº TUA`, mostrando os
parâmetros, as obrigações de amostragem/tratamento e todos os critérios a
respeitar.

Tudo o que é preciso já existe em `Power BI Data\` — só falta voltar a ligar
duas tabelas ao modelo e desenhar a página. Faz-se inteiramente no Power BI
Desktop, sem tocar nos ficheiros `.tmdl` à mão.

---

## Parte A — Trazer as tabelas Conditions e Autocontrolo para o modelo

Foram removidas quando reduzimos o modelo à Visão Geral. Voltam a entrar assim:

1. Separador **Base** → **Obter dados** → **Excel (pasta de trabalho)**.
2. Navegar até `…\EPAL-project\Power BI Data\Conditions.xlsx` → **Abrir**.
3. No **Navegador**, marcar a folha **Conditions** → **Transformar dados**
   (não "Carregar") para abrir o Power Query. Marcar **"Utilizar a primeira
   linha como cabeçalhos"** se ainda não estiver.
4. No Power Query, acertar os tipos das colunas numéricas (senão dão erro ou
   ficam com `.0`):
   - `VLE mín`, `VLE máx`, `Carga máx. admissível (kg/dia)` → clicar no ícone de
     tipo do cabeçalho → **Número decimal**. Se der erro de conversão, usar
     **Usando definições regionais…** → tipo *Número decimal*, localidade
     *Inglês (Estados Unidos)* (os pontos decimais são ingleses).
   - As colunas de critério (`≤ 100% VLE (dobro)`, `≤ 150% VLE`,
     `≤ uma ordem de grandeza do VLE`, `média mensal ≤ VLE`, `média anual ≤ VLE`,
     `Quadro III DL 152/97 (borlas)`, `Gama de valores (intervalo)`) → deixar
     como **Número inteiro** (contêm 0/1) ou **Texto** — tanto faz.
5. Renomear a consulta (painel direito, **Nome**) para `Conditions`.
6. Repetir os passos 1–5 para `Autocontrolo.xlsx` (folha `Autocontrolo`) →
   nome `Autocontrolo`. Aqui a
   coluna numérica a acertar é `Nº análises requeridas` → **Número inteiro**.
7. **Base → Fechar e aplicar**.

> Dica: para que estas consultas usem a mesma pasta que as outras, no Editor
> Avançado de cada uma troca o caminho fixo por
> `Excel.Workbook(File.Contents(DataFolder & "\Conditions.xlsx"), null, true)`
> seguido de `Source{[Item="Conditions",Kind="Sheet"]}[Data]` — tal como a
> consulta `Licenses` já faz. Mantém tudo a apontar para o parâmetro `DataFolder`.

---

## Parte B — Criar as relações (esquema em estrela)

1. Ir à vista **Modelo** (ícone de tabelas ligadas, à esquerda).
2. Arrastar `Conditions[Nº TUA]` para cima de `Licenses[Nº TUA]`.
   - Deve criar uma relação **Muitos-para-um (\*:1)**, sentido **único**,
     de Conditions → Licenses. Se o Power BI propuser outra coisa, faz duplo
     clique na relação e confirma: *Cardinalidade = Muitos para um*,
     *Direção do filtro cruzado = Única*.
3. Arrastar `Autocontrolo[Nº TUA]` para `Licenses[Nº TUA]` — mesma coisa.

Resultado: `Licenses` ao centro, com `Dim_Estado`, `Conditions` e `Autocontrolo`
à volta. `Licenses` filtra as outras, nunca ao contrário.

---

## Parte C — Transformar uma página nova em alvo de drill-through

1. Em baixo, no separador de páginas, clicar **+** para nova página. Duplo
   clique no nome → **Ficha da Licença**.
2. Com a página vazia selecionada, no painel **Visualizações** procurar a caixa
   **Drill-through** (Fazer drill-through). Arrastar `Licenses[Nº TUA]` para
   **"Adicionar campos de drill-through aqui"**.
   - É este passo que torna a página um destino de drill-through: passa a abrir
     filtrada pelo `Nº TUA` de onde vieres.
3. Deixar **"Manter todos os filtros"** ligado (opcional).
4. Repara que aparece automaticamente um **botão de seta para trás** no canto —
   é o "voltar" para a Visão Geral. Posiciona-o onde quiseres.
5. Recomendado: clicar com o botão direito no separador **Ficha da Licença** →
   **Ocultar página**, para que só se abra por drill-through e não apareça na
   barra de páginas.

---

## Parte D — Visuais da Ficha da Licença

### 1. Cabeçalho — identidade e estado da licença
Inserir um **Cartão de várias linhas** (multi-row card) ou vários **Cartões**, e
arrastar de `Licenses`:
`Estabelecimento`, `Nº TUA`, `Estado da Licença`, `Data de Validade`,
`Nível de tratamento`, `Esquema de tratamento`.
> Como a página estará filtrada por um único `Nº TUA`, cada cartão mostra o valor
> dessa licença. O `Nível/Esquema de tratamento` cobre a parte "tratamentos".

### 2. Tabela — Parâmetros e VLE
Inserir uma **Tabela** e, de `Conditions`, arrastar por esta ordem:
`Parâmetro`, `VLE`, `VLE mín`, `VLE máx`, `Carga máx. admissível (kg/dia)`,
`Frequência de amostragem`, `Tipo de amostragem`, `Legislação aplicável`.
> Estes são os limites e a periodicidade que a licença impõe a cada parâmetro.

### 3. Tabela — Obrigações de autocontrolo (colheita de amostras)
Inserir outra **Tabela** e, de `Autocontrolo`, arrastar:
`Local de amostragem`, `Parâmetro`, `Frequência de amostragem`,
`Tipo de amostragem`, `Nº análises requeridas`.
> É o plano de monitorização: onde, o quê, com que frequência e quantas análises
> por ano.

### 4. Matriz — Critérios a respeitar por parâmetro
Inserir uma **Matriz**:
- **Linhas** → `Conditions[Parâmetro]`.
- **Valores** → as 7 colunas de critério de `Conditions`:
  `média mensal ≤ VLE`, `média anual ≤ VLE`, `≤ 100% VLE (dobro)`,
  `≤ 150% VLE`, `≤ uma ordem de grandeza do VLE`, `Gama de valores (intervalo)`,
  `Quadro III DL 152/97 (borlas)`.
- Cada valor é 0/1. Para cada um, clicar na seta do campo → **Não resumir**
  (ou *Máximo*), para não somar.

Para mostrar um ✓ em vez de 1:
- Selecionar a matriz → **Formatar** (pincel) → **Elementos da célula** →
  ligar **Ícones** → **Formatação condicional** na coluna → regra:
  *se valor = 1 → ✓ verde; senão → em branco*.
- Alternativa mais simples: em vez da matriz de flags, usar uma tabela com
  `Parâmetro` + `Avaliação da Conformidade Legal` (o texto legal completo do
  critério) + `Estado do critério`.

---

## Parte E — Fazer o clique funcionar a partir da Visão Geral

Com o campo de drill-through definido (Parte C), já funciona:

1. Na página **Visão Geral**, clicar com o **botão direito** numa linha da
   tabela de licenças (ou numa bolha do mapa) → **Fazer drill-through** →
   **Ficha da Licença**.
2. A Ficha abre filtrada por essa licença. O botão de seta traz de volta.

Opcional — botão sempre visível em vez do menu de contexto:
1. **Inserir → Botões → Em branco**. Texto: "Ver ficha da licença".
2. Selecionar o botão → **Formatar botão → Ação → Tipo = Drill-through →
   Destino = Ficha da Licença**.
3. O botão só fica ativo quando há uma licença selecionada na tabela.

---

## Parte F — Verificação final

- Atualizar (**Base → Atualizar**) para importar Conditions e Autocontrolo.
- Testar 2–3 licenças, incluindo uma **LURH** e uma **TUA**, e uma **caducada**.
- Confirmar que a matriz de critérios muda de parâmetros entre licenças.
- Se uma licença não tiver linhas de Autocontrolo, a tabela fica vazia — normal
  (nem todas têm plano de autocontrolo no documento).

---

## Notas específicas deste projeto

- **Fontes das colunas**: parâmetros, VLE e critérios vêm de `Conditions.xlsx`
  (1207 linhas); a colheita de amostras vem de `Autocontrolo.xlsx` (953 linhas);
  o tratamento está em `Licenses`.
- **TUA e LURH juntos**: ambas as origens já estão nestes ficheiros (coluna
  `TUA/LURH` em Licenses), por isso a Ficha funciona igual para as duas.
- **Não uses o `build_geral_page.py`** depois disto: a página Visão Geral está a
  ser mantida à mão; o script recriaria os visuais.
- **Se quiseres que eu construa a página por ti** em vez de a fazeres à mão,
  diz — consigo gerar a página e as relações diretamente nos ficheiros do
  projeto.
