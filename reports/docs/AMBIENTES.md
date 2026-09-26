# Onde é que este pipeline corre, afinal?

**Estado: por confirmar.** Este ficheiro regista o que os registos mostram, não
o que alguém decidiu. Se souber a resposta, corrija-o — vale mais uma linha
certa aqui do que a arqueologia que foi precisa para escrever isto.

## O ambiente esperado

Windows, com o Python portátil da própria pasta (`runtime\python`), lançado
pelos `.bat`. É isto que o `LEIA-ME_PRIMEIRO.txt` descreve e é para isto que
tudo está afinado.

## O ambiente que apareceu nos registos

Durante a auditoria de 24-08-2026 encontraram-se execuções feitas **fora do
Windows**, num sistema de ficheiros Linux, com o projeto montado em
`/mnt/EPAL`:

| Registo | Sessão | Linhas |
|---|---|---|
| `Extraction_Code_TUA/data/logs/tua_20260810.log` | `sleepy-blissful-clarke` | 1 |
| `Extraction_Code_TUA/data/logs/tua_20260823.log` | `rcw-01f1dmyunqud94ur1soshkxh` | 3 |
| `Extraction_Code_LURH/data/logs/lurh_20260823.log` | `rcw-01f1dmyunqud94ur1soshkxh` | 24 |

Duas sessões distintas, a mais recente na véspera da auditoria. Duas licenças
LURH ficaram em `data/review` por causa disso, com uma nota que aponta para um
caminho `/sessions/...` que ninguém neste projeto reconhece:

```
Extraction crashed: Failed to open file
'/sessions/rcw-01f1dmyunqud94ur1soshkxh/mnt/EPAL/EPAL-project/Drop_New_Licenses/...'
```

## Porque é que isto importa

Não é que correr em Linux seja errado — é que **ninguém sabia que acontecia**,
por isso nada foi verificado nesse ambiente. As diferenças que já se
conhecem:

- **Maiúsculas nos nomes.** `glob('*.pdf')` é indiferente a maiúsculas no
  Windows e sensível no Linux. Um ficheiro `LICENCA.PDF` era processado num e
  ignorado **em silêncio** no outro. Corrigido: os dois pipelines usam agora
  `_inbox_pdfs()`, que trata os dois casos igual.
- **A ligação estável do Power BI** (`C:\EPAL\powerbi`, ver
  `fix_powerbi_datafolder.py`) usa `mklink /J`. Fora do Windows não existe — o
  script avisa e usa o caminho da máquina.
- **`make_deploy.py`** já avisava que um pacote gerado fora do Windows leva um
  `DataFolder` inválido.
- **`os.startfile()`**, que abre o Excel no fim de cada lote, só existe no
  Windows. Está protegido por `try/except`, portanto não parte nada.

## O que é preciso decidir

1. Aquelas execuções foram deliberadas (um ambiente de teste na nuvem?) ou
   acidentais?
2. Se forem para continuar: esse ambiente passa a ser suportado e testado, e
   este ficheiro deixa de dizer "por confirmar".
3. Se não: as duas licenças em `data/review` que falharam por causa disto
   devem ser reprocessadas no Windows. São elas:
   - `Benavila L006744.2022.RH5A (11.10.2022 a 10.10.2027).pdf`
   - `São Vicente de Valongo_L001763.2021.RH7 (25.01.2021 a 24.01.2031).pdf`
