# Botão "Abrir licença" — abrir o PDF a partir do dashboard

Põe um botão no relatório que abre o PDF da licença que está a ser vista.

O Power BI **não** consegue abrir um ficheiro a partir de um caminho qualquer:
precisa de um endereço numa coluna marcada como *Web URL*. O feed já traz essa
coluna — `Ficheiro da Licença`, na tabela `Licences` — construída pelo
`make_powerbi_all.py`.

---

## Onde a coluna aponta (e como mudar depois da migração)

**Agora:** endereços `file:///` para a pasta deste computador. Abrem no Power
BI **Desktop**. Não servem para o relatório publicado nem para outra pessoa —
levam o caminho de quem gerou o feed.

**Depois de a pasta estar no SharePoint**, é uma linha em
`src/Common_Code_PowerBI/pdf_link.py`:

```python
BASE_PADRAO = BASE_SHAREPOINT
```

…e voltar a correr `make_powerbi_all.py`. O endereço do SharePoint já está lá
guardado em `BASE_SHAREPOINT`. **Confirmar primeiro** abrindo um PDF qualquer
no SharePoint pelo browser e comparando — sobretudo se a pasta `EPAL-project`
existe mesmo lá dentro, ou se o conteúdo foi copiado directamente para
`Licences project`.

Para experimentar sem mexer no código: `EPAL_PDF_BASE_URL=<endereço>`.
`EPAL_PDF_BASE_URL=-` volta ao modo local.

---

## Já está feito no modelo

Os passos 1 e 2 (marcar a coluna como *URL da Web* e criar as medidas) **já
estão no ficheiro do modelo**. Ao abrir o projeto encontra:

* a coluna `Ficheiro da Licença` na tabela `Licences`, com
  `dataCategory: WebUrl`;
* a medida **`URL da Licença`** — devolve o endereço só quando está UMA
  licença seleccionada;
* a medida **`Dica do botão Licença`** — explica porque é que o botão está
  cinzento.

Falta apenas pôr o botão na página, que é coisa de meio minuto:

1. **Inserir → Botões → Em branco**.
2. Formatação: **Estilo → Texto** → `Abrir licença`.
3. **Ação** → **Ativado**, **Tipo** = **URL da Web**.
4. Ao lado de *URL da Web*, clicar em **fx** → medida **`URL da Licença`**.
5. (recomendado) **Estilo → Dica de ferramenta → fx** → medida
   **`Dica do botão Licença`**.

O botão desliga-se sozinho quando não há exactamente uma licença escolhida.

---

## O que esperar

**As 700 licenças do feed têm ligação** para o seu PDF. Se alguma vez alguma
ficar sem, é porque o `file_name` no master deixou de corresponder a um PDF do
arquivo — nesse caso o botão fica cinzento e a dica explica porquê, em vez de
abrir um endereço partido.

Para saber quais são, a qualquer momento:

```
runtime\python\python.exe src\_verificar_cobertura.py --listar
```

## Se o botão não abrir nada

| Sintoma | Causa provável |
|---|---|
| Botão sempre cinzento | Há zero ou mais de uma licença seleccionada (é o comportamento certo) |
| Funciona no Desktop, não no Service | Normal por agora: a coluna tem `file:///`. Ver a secção do endereço, acima |
| Abre e dá "não encontrado" | O PDF foi movido ou renomeado depois do último `make_powerbi_all.py` |
| Pede credenciais | Normal na primeira vez no browser: é o SharePoint a autenticar |
