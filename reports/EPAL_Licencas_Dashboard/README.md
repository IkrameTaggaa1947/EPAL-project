# EPAL Licenças Dashboard (TUA + LURH)

Relatório de página única — **Visão Geral** — sobre as licenças TUA e LURH no
mesmo modelo, com a paleta EPAL (azul `#0D6ABF`, laranja `#FF911B`).

## Dados
O modelo lê os ficheiros **.xlsx** combinados em `Power BI Data/` (parâmetro
`DataFolder`; uma folha por ficheiro, com o mesmo nome do ficheiro).
Regenerar sempre que qualquer um dos pipelines correr:

    python extraction/make_powerbi_all.py

A coluna `TUA/LURH` distingue a origem das 269 licenças.

### TabelaCentroide_ETAR (registo de ativos)

Dimensão ligada a `Licenses` por `Chave_ETAR` — o nome da ETAR normalizado
(sem acentos, sem artigos *de/da/do*), calculado em `src/etar_key.py` e escrito
nos dois ficheiros por `make_powerbi_all.py`. A licença não traz nenhum código
de ativo, por isso o nome é a única ligação possível.

* Relação: `Licenses[Chave_ETAR]` (*) → `TabelaCentroide_ETAR[Chave_ETAR]` (1),
  filtro no sentido normal — o segmentador do Centro Operacional filtra tudo.
* **227 das 269 licenças** têm correspondência. As outras 42 ficam com
  `Chave_ETAR` vazia e **desaparecem quando um Centro Operacional é
  selecionado** — é o comportamento pedido, não um erro. Para as recuperar,
  acrescentar o par nome-da-licença → nome-do-registo a `OVERRIDES` em
  `src/etar_key.py` e voltar a correr `make_powerbi_all.py`.
* A tabela traz também `CODMAXIMO`, o mesmo código de
  `WWTP_History[CodigoMaximo]` — é por aqui que as licenças se ligam ao
  histórico laboratorial do WWTP_Report.

## Abrir

1. **Primeiro passo em cada máquina: corrigir o `DataFolder`.** É um caminho
   **absoluto** — o Power Query não tem caminhos relativos ao `.pbip`, por isso o
   valor guardado no git é o da máquina de quem fez o último commit. Abrir
   **Transformar dados → Gerir parâmetros** e apontar `DataFolder` para
   `<raiz do projeto>\data\powerbi`.
   Se for esquecido, a atualização falha com um erro do Power Query que nomeia um
   ficheiro `.xlsx` e nunca o parâmetro. Para confirmar antes de abrir o Desktop:

       python reports/validate_project.py

2. Abrir `EPAL_Licencas_Dashboard.pbip` no Power BI Desktop.
3. Atualizar.

Os ficheiros do feed estão versionados, por isso um clone do repositório já os
tem — o `make_powerbi_all.py` só é preciso depois de correr as extrações.

## Página

**Visão Geral** (única página do relatório)

| Zona | Conteúdo |
|------|----------|
| Cartões | Total · Caducadas · A expirar (< 8 meses) · Válidas · Sem data |
| Mapa | Localização das ETAR, cor = estado de validade |
| Circular | Distribuição por estado de validade |
| Colunas | Expirações por ano de validade, empilhadas por estado |
| Barras | Licenças por região |
| Tabela | Todas as licenças, ordenadas por data de validade (as que expiram primeiro no topo) |
| Filtros | Tipo (TUA/LURH) · **Centro Operacional** (pendente, no cabeçalho) |

Colunas da tabela: ETAR · Tipo · Nº do título · Região · Município · Estado ·
Validade · Dias até expirar · Tratamento. Para mudar as colunas, editar a lista
do visual `tabela_licencas` em `powerbi/build_geral_page.py` e voltar a correr o
script. `Dias até expirar` é a medida `Medidas[Dias até Expirar]` (negativo =
já caducada).

Classificação de expiração: **Caducada** (validade anterior a hoje) ·
**A expirar** (expira dentro de 8 meses) · **Válida** (mais de 8 meses) ·
**Sem data** (o documento não indica validade — 5 TUA dizem "Sem dados").

O limiar de 8 meses está definido **uma única vez**, na medida
`Medidas[Meses de Antecedência]`. Alterar aí muda todo o relatório.

## Modelo — esquema em estrela (só o necessário para a Visão Geral)

```
  Dim_Estado ─►  Licenses  ◄─  (Medidas: tabela de medidas, sem dados)
```

Três tabelas apenas:

- `Licenses` — dimensão de licença (geografia REGIÃO/ARH/Município como
  dimensão degenerada). Fonte: `Power BI Data/Licenses.csv`.
- `Dim_Estado` — os 4 estados de validade (Caducada / A expirar < 8 meses /
  Válida / Sem data), com `Ordem` e `Cor`.
- `Medidas` — todas as métricas.

As tabelas `Conditions`, `Autocontrolo`, `Legislacao`, `Avaliacao` e a tabela de
datas `Calendario` foram **removidas**: nenhum visual da página as usava. Se
mais tarde precisar de páginas de VLE / autocontrolo / calendário, recupere-as
do git (commit `7b81f75`) e volte a adicionar as respetivas `ref table` +
relações.

Uma única relação: `Licenses[Estado da Licença]` 1—* `Dim_Estado[Estado]`.


## Regenerar a página

O layout é gerado por script, para se manter consistente:

    python powerbi/build_geral_page.py            # reescreve section0/visuals
    python powerbi/build_geral_page.py --check    # valida referências ao modelo

O `--check` confirma que todos os campos usados nos visuais existem no modelo
semântico — correr sempre depois de mexer no modelo.

## Nota para quem editar o TMDL à mão

O parser TMDL do Power BI **não aceita comentários**: nem `//`, nem `/* */`, nem
blocos `///` soltos. Só admite descrições `///` imediatamente coladas a uma
declaração (`table`, `column`, `measure`). Qualquer outra coisa faz o projeto
recusar-se a abrir com:

    Erreur de format TMDL : InvalidLineType — Other! / Empty!

Por isso a documentação do modelo vive neste README e não dentro dos ficheiros
`.tmdl`. Antes de abrir o projeto no Power BI Desktop:

    python powerbi/validate_project.py

que verifica isto, além de: indentação por tabulações, code fences, partições M
truncadas, colunas das relações, `ref table` vs ficheiros, ligações linguísticas
das culturas, JSON dos visuais e campos usados pelos visuais.

## Tipos de dados nos CSV

`extraction/make_powerbi_all.py` normaliza as colunas numéricas de
`Licenses.csv` e `Autocontrolo.csv` antes de as escrever (`normalise_numbers`):
inteiros sem parte decimal, decimais com **vírgula**. Sem isso, os valores LURH
lidos via pandas saíam como `"2002.0"` (falha a conversão para número inteiro,
erro na linha inteira) e as coordenadas saíam com **ponto** enquanto as TUA
usavam vírgula — e como a consulta lia a coluna em `pt-PT`, onde o ponto é
separador de milhares, `-7.536845` tornava-se `-7536845` e atirava o ponto do
mapa para fora do país.

As consultas `Licenses` e `Autocontrolo` passaram também a converter números de
forma tolerante (aceitam vírgula ou ponto, e `12.0` num campo inteiro), pelo que
um formato inesperado deixa a célula vazia em vez de rebentar com a linha.

## Histórico

- **2026-07-21** — Reposto para uma única página (Visão Geral); removidas as
  páginas Renovações, Ficha da Licença, Condições VLE e Autocontrolo.
  Modelo reestruturado em estrela (novas `Dim_Estado` e `Calendario`), lógica
  movida para medidas, limiar de expiração passado de 6 para 8 meses.
  Corrigido `Licenses.tmdl`, cuja partição M estava **truncada a meio da
  expressão** (faltavam o `in Inteiros` e a anotação final) — o modelo não abria.
  Removidas ligações linguísticas em `cultures/pt-PT.tmdl` que apontavam para
  colunas entretanto eliminadas. Versão anterior (5 páginas) preservada no git,
  commit `7b81f75`.
- **2026-07-21** — Adicionada a tabela de todas as licenças (página passa a
  1280x964) e as medidas `Dias até Expirar` e `Estado (texto)`.
- **2026-07-21** — Corrigidos os erros de tipo em `Licenses` (140 linhas LURH em
  erro: `Ano de arranque` = "2002.0") e o separador decimal das coordenadas LURH.
  Mesmo defeito preventivamente corrigido em `Autocontrolo`
  (`Nº análises requeridas`, 938 linhas).
