# Guia — Configurar o projeto num computador novo da equipa

Para quem está a levar o projeto a um novo colega (não para o colega ler
sozinho — para isso já existe o `LEIA-ME_PRIMEIRO.txt` na raiz do projeto).
Este guia serve para confirmar que a instalação correu bem e para decidir,
como equipa, quem liga o vigia automático.

---

## Antes de começar

Confirme que a pasta `EPAL-project` está **totalmente sincronizada** no
OneDrive desse computador (ícone verde, sem nuvens azuis pendentes). Se
ainda estiver a sincronizar, os passos abaixo podem falhar de forma
confusa — não é um erro do projeto, é só esperar.

## Passo 1 — Correr o configurador

Pedir ao colega para fazer duplo-clique em:

    CONFIGURAR_ESTE_PC.bat

Não precisa de direitos de administrador e não instala nada — só confirma
que encontra o Python já incluído na pasta partilhada e escreve as
definições deste PC.

## Passo 2 — Ler o diagnóstico final

O `.bat` termina a mostrar um diagnóstico. O que deve aparecer:

| Linha | O que deve dizer |
|---|---|
| Python | `3.13.1` e o caminho deve terminar em `runtime\python\python.exe` |
| ROOT, data/powerbi, Drop TUA, Drop LURH, reports, outputs | `[OK]` em todas |
| epal.local.ini | `[OK]` |
| venv local | `[EM FALTA]` — **é normal**, o runtime portátil tem prioridade e o venv local nunca chega a ser usado |
| Power BI DataFolder | `-> CORRETO` |

Se **Python** disser "AVISO: o Python portátil ainda não existe nesta
pasta" e passar a usar o Python do computador (ou nenhum), a causa mais
provável é a pasta ainda não ter sincronizado por completo — voltar a
correr o `.bat` passados uns minutos resolve quase sempre. Só volte a
correr o `INSTALAR_PYTHON_PORTATIL.bat` se isto persistir depois de a
sincronização terminar (esse script só devia mesmo correr uma vez, no
computador original).

Qualquer outra linha que não diga `[OK]` — enviar um print deste ecrã
antes de continuar.

## Passo 3 — NÃO testar com um PDF real ainda

O `CONFIGURAR_ESTE_PC.bat` já é o teste completo — não mexe em nenhum
PDF nem nos ficheiros master. Não é preciso (nem aconselhável) largar um
PDF de teste só para confirmar a instalação.

## Passo 4 — Decidir quem liga o vigia automático

**Só UM computador da equipa pode ter o vigia automático
(`register_watch_task.bat`) ligado.** A pasta é partilhada — se dois
computadores processarem os mesmos PDFs ao mesmo tempo, os ficheiros
master podem ficar corrompidos.

Isto já tinha um aviso escrito, mas nada o impedia tecnicamente até agora.
Desde 2026-08-24, o `watch_all.py` tem uma proteção adicional: cada ciclo
regista qual o computador ativo num ficheiro partilhado
(`src/automation_of_extraction/data/watch_lock.json`) e um segundo
computador que tente correr ao mesmo tempo é automaticamente ignorado (não
mexe em nada, fica só um aviso no registo) — a não ser que o primeiro
computador esteja parado há mais de 10 minutos, caso em que o novo assume
sozinho. **Isto é uma rede de segurança extra, não substitui o acordo da
equipa** — continuem a decidir explicitamente qual É o computador do
vigia.

Para ver a qualquer momento qual computador está ativo, correr (usando o
Python do projeto):

    runtime\python\python.exe src\automation_of_extraction\watch_all.py --status

Nos restantes computadores, usar sempre `PROCESSAR_AGORA.bat`
manualmente. Se um dia quiserem mudar qual é o computador do vigia, correr
`unregister_watch_task.bat` no antigo primeiro — liberta a marca logo, sem
esperar os 10 minutos.

## Se algo correr mal

Mesma lista do `LEIA-ME_PRIMEIRO.txt` — reproduzida aqui por conveniência:

- **Power BI diz que não encontra os dados** → fechar o Power BI e
  duplo-clique em `ATUALIZAR_POWERBI.bat`.
- **"NÃO FOI ENCONTRADO NENHUM PYTHON"** → a pasta ainda pode estar a
  sincronizar; esperar e tentar de novo.
- **Uma licença foi para `data\review`** → problema a ler esse PDF
  específico, não a instalação; há um `.txt` ao lado a explicar porquê.
- Para pedir ajuda, pedir ao colega para correr e enviar o texto que
  aparece:

      runtime\python\python.exe src\epal_config.py

---

*Nota (2026-08-24): a instalação foi confirmada a funcionar no primeiro
computador (o do Ikrame). Este guia ainda não foi testado no processo de
adicionar um SEGUNDO computador — ao usá-lo pela primeira vez, vale a pena
anotar qualquer passo que não corresponda ao que aqui está descrito.*
