# Pipeline LURH/TUA — Guia de instalação e manutenção local

Sistema 100% local e determinístico para extrair condições de licenças de
rejeição (TUA/LURH em PDF do SILIAMB), consolidar no ficheiro mestre e manter
o ficheiro de critérios de conformidade sincronizado. Não precisa de internet
(exceto o fallback LLM opcional), servidor ou base de dados.

---

## 1. Instalação numa máquina nova

1. Instalar **Python 3.10 ou superior** (python.org — marcar "Add to PATH").
2. Copiar a pasta do projeto completa para o disco local, p.ex. `C:\LURH\`.
3. Numa linha de comandos dentro da pasta:

   ```
   pip install -r requirements.txt
   ```

4. Criar as pastas de dados (se não existirem): `data\inbox`,
   `data\review`, `data\logs` (as licencas boas vao para `..\..\data\pdfs\NEW\LURH`).
5. Testar com um PDF: colocá-lo em `data\inbox` e correr `PROCESSAR_AGORA.bat`.

## 2. Ficheiros do sistema

| Ficheiro | Papel |
|---|---|
| `pipeline.py` | Orquestrador: extrai → QA → atualiza mestre → sincroniza critérios |
| `extract_tua.py` | Leitura do PDF (PyMuPDF, tabelas por cabeçalho/marcador) |
| `qa.py` | Verificações de qualidade → OK / REVIEW / FAILED |
| `enrich.py` | Atualização do `master_lurh.xlsx` + export CSV para Power BI |
| `sync_criterios.py` | Sincroniza novas "Avaliações de Conformidade" com o ficheiro de critérios |
| `criterios_rules.json` | Regras aprendidas (criado automaticamente; faz parte do conhecimento do sistema) |
| `PROCESSAR_AGORA.bat` | Execução manual |
| `watch_all.bat` + `Automation files\register_watch_task.bat` | Execução automática (Tarefa Agendada do Windows) |
| `master_lurh.xlsx` | Ficheiro mestre consolidado |
| `CriteriosConformidade_geral_*.xlsx` | Ficheiro de critérios (folhas AdVT_CriteriosConformidade e CriteriosConformidade_GLOBAL) |

## 3. Operação diária

- **Manual:** PDF em `data\inbox` → duplo clique em `PROCESSAR_AGORA.bat`.
- **Automática:** correr `Automation files\register_watch_task.bat` uma vez (como administrador)
  para registar a tarefa agendada; `Automation files\unregister_watch_task.bat` remove-a.
- Resultado de cada execução no log `data\logs\tua_AAAAMMDD.log`.
- PDFs com problemas vão para `data\review` — nunca são ignorados em silêncio.

### Regras de ouro
- O ficheiro de critérios e o mestre devem estar **fechados no Excel** durante
  a execução (o Excel bloqueia o ficheiro; se acontecer, o sync grava uma
  cópia `_PENDING_` e avisa no log).
- Antes de gravar, o sync cria sempre uma cópia `.bak_AAAAMMDD_HHMMSS`.

## 4. Ciclo de revisão humana (obrigatório)

Após cada lote, filtrar a última coluna de `AdVT_CriteriosConformidade`:

| Marca | Origem | Ação |
|---|---|---|
| `AUTO — rever` | Regras de palavras-chave | Confirmar critérios; substituir a marca por iniciais/data |
| `AUTO(regra aprendida) — rever` | Regra aprendida (JSON) | Confirmação rápida |
| `AUTO-LLM — rever` | Classificação por LLM (nova lógica) | Rever com mais cuidado |
| `AUTO(sem classificação) — rever` | Nada correspondeu (sem API) | Classificar manualmente; se recorrente, acrescentar regra (ver 5) |

**Nunca apagar o texto da Avaliação** — é a chave de deduplicação.

## 5. Manutenção da lógica

- **Nova lógica legal recorrente:** acrescentar uma linha ao dicionário em
  `classify()` de `sync_criterios.py` (ex.: `'C8_percentil': 'percentil 95' in t`)
  e a célula correspondente no bloco de escrita. Uma alteração, dois sítios.
- **Layout do PDF mudou (SILIAMB):** ajustar os marcadores/cabeçalhos em
  `extract_tua.py`.
- **Fallback LLM (opcional, requer internet):** definir a variável de ambiente
  `ANTHROPIC_API_KEY` (`setx ANTHROPIC_API_KEY "sk-ant-..."`). Sem a chave o
  sistema funciona igual, apenas sem classificação automática de lógica nova.
- **Teste sem gravar:** `python sync_criterios.py --dry-run`.

## 6. Cópias de segurança

Fazer backup regular (copiar para outra pasta ou disco) de:
`master_lurh.xlsx`, `CriteriosConformidade_geral_*.xlsx`,
`criterios_rules.json`, todos os `.py`/`.bat` e `requirements.txt`.
A pasta `data\logs` é a auditoria; o arquivo dos PDFs é `..\..\data\pdfs`.

Para reinstalar noutra máquina: passo 1–3 desta página + restaurar o backup.

## 7. Resolução de problemas

| Sintoma | Causa provável | Solução |
|---|---|---|
| `PermissionError` / ficheiro `_PENDING_` | Excel aberto | Fechar Excel e repetir |
| Linhas novas "invisíveis" | Filtro antigo ativo | O sync estende o filtro; limpar filtros manuais |
| Mesmo texto readicionado | Texto da Avaliação foi editado na folha | Editar só as colunas de critério, nunca o texto |
| `Master ... missing columns` | Cabeçalhos do mestre alterados | Alinhar nomes em `master_triples()` |
| PDF em `review/` | QA detetou anomalia | Ver log, corrigir/validar manualmente |
