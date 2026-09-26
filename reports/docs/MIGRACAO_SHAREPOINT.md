# Migrar do OneDrive pessoal para o SharePoint da equipa

Ler **antes** de criar a biblioteca. Uma das decisões — o nome do site e da
biblioteca — não se pode corrigir depois sem voltar a sincronizar tudo.

Depois de migrar, cada pessoa corre isto no seu PC e deve dar `TUDO PRONTO`:

```
runtime\python\python.exe src\_verificar_migracao.py
```

---

## 1. O nome da biblioteca decide se isto funciona

Este é o único problema que falha **em silêncio**: os ficheiros não aparecem,
e não há mensagem de erro nenhuma.

O Windows corta os caminhos aos **260 caracteres**. Dentro do projeto, o
caminho mais comprido tem **146** — é uma pasta de visual do Power BI, com
nome gerado automaticamente, que não se pode encurtar.

```
260  limite do Windows
-146 caminho mais longo dentro do projeto
-  1 separador
────
 113  ← máximo para a pasta onde o SharePoint sincroniza
```

Hoje, no OneDrive pessoal, a pasta tem **95** caracteres: sobram 18. Um
caminho de equipa é quase sempre **mais comprido**, porque tem o nome da
organização, do site e da biblioteca:

```
C:\Users\<nome>\<Organização>\<Site> - <Biblioteca>\EPAL-project
```

### O destino escolhido: `2026-InternshipatEPAL`

A pasta escolhida foi:

```
/sites/2026-InternshipatEPAL/Documentos Partilhados/General/Deliverables/Licences project
```

São **três níveis** dentro da biblioteca — `General\Deliverables\Licences
project` — e é isso que estoira o orçamento. Sincronizando a biblioteca da
forma normal, o caminho local fica:

```
C:\Users\Ikrame TAGGAA\<Org>\2026-InternshipatEPAL - Documentos Partilhados\General\Deliverables\Licences project\EPAL-project
```

| `<Org>` no Explorador | Total | Resultado |
|---|---|---|
| `EPAL` | 125 | **12 acima do limite** |
| `PrivEPAL` | 129 | **16 acima** |
| nome completo da empresa | 163 | **50 acima** |

**Não cabe em nenhum dos casos.** Ficheiros do Power BI iriam faltar, sem
aviso nenhum.

### A solução: "Adicionar atalho ao OneDrive"

Em vez de sincronizar a biblioteca inteira, abrir a pasta **`Licences
project`** no browser e usar **"Adicionar atalho ao OneDrive"**. Os níveis
intermédios desaparecem do caminho local:

```
C:\Users\Ikrame TAGGAA\<Org>\Licences project\EPAL-project
```

| `<Org>` no Explorador | Total | Margem |
|---|---|---|
| `EPAL` | 57 | 56 caracteres |
| `PrivEPAL` | 61 | 52 |
| nome completo da empresa | 95 | 18 |

Cabe em todos os casos. **É esta a forma de sincronizar**: cada pessoa faz o
atalho na pasta `Licences project`, nunca na biblioteca inteira.

> O nome exacto de `<Org>` é o que aparece no painel esquerdo do Explorador.
> O verificador (`src\_verificar_migracao.py`) diz o número real em cada PC —
> é esse que conta, não estas estimativas.

---

## 2. O `.git` não pode ser sincronizado

O repositório tem 989 ficheiros e o git reescreve o índice e ficheiros de
bloqueio a toda a hora. Um cliente de sincronização a mexer nisso ao mesmo
tempo **corrompe o repositório** — é uma das formas mais conhecidas de perder
um histórico.

Duas saídas, ambas boas:

**A. Um remoto a sério (recomendado).** Criem um repositório no Azure DevOps
ou GitHub da organização, façam `git push`, e o histórico deixa de depender da
pasta partilhada:

```bash
git remote add origin <url-do-repositorio>
```

```bash
git push -u origin main
```

**B. Manter o git fora da pasta sincronizada.** Trabalhem num clone local
(por exemplo `C:\EPAL\src`) e usem a pasta do SharePoint só para os dados e
para quem executa. Mais simples de explicar, mas obriga a copiar código à mão.

Enquanto nada disto estiver feito, o `.git` continua dentro da pasta. Não é
imediato — mas é uma questão de tempo.

---

## 3. O que fazer, por ordem

1. **Cada pessoa faz "Adicionar atalho ao OneDrive" na pasta `Licences
   project`** — não sincronizar a biblioteca inteira (ver ponto 1). É este
   passo que decide se os ficheiros chegam todos.
2. **Copiar a pasta `EPAL-project` inteira** para lá. Usar o *Explorador de
   Ficheiros* depois de criar o atalho, ou o *Migration Manager* da
   Microsoft. Não usar arrastar-e-largar no browser: perde ficheiros com
   caminhos longos sem avisar.
3. **Esperar que a sincronização termine mesmo.** O ícone do OneDrive tem de
   ficar verde e parado. São ~11 000 ficheiros.
4. **Cada pessoa corre `Automation files\CONFIGURAR_ESTE_PC.bat`.** É obrigatório outra vez:
   os ficheiros são novos, portanto o Python volta a ser marcador na nuvem, e
   a ligação `C:\EPAL\powerbi` ainda aponta para a pasta antiga.
5. **Cada pessoa corre o verificador** e confirma `TUDO PRONTO`:
   ```
   runtime\python\python.exe src\_verificar_migracao.py
   ```
6. **Escolher UM computador para o vigia automático**
   (`Automation files\register_watch_task.bat`). Continua a valer o mesmo aviso: dois PCs a
   processar os mesmos PDFs corrompem os ficheiros master.
7. **Só depois** apagar a cópia antiga do OneDrive pessoal.

---

## 4. O que muda para cada pessoa

| | Antes | Depois |
|---|---|---|
| Pasta | OneDrive pessoal de uma pessoa | Biblioteca da equipa |
| Quem tem os dados | uma pessoa | quem tiver acesso à biblioteca |
| Espaço em disco | ~1 GB | ~1 GB **em cada PC que sincronize** |
| `C:\EPAL\powerbi` | aponta para o OneDrive | tem de ser recriada |
| Python fixado | sim, neste PC | não — voltar a fixar em cada PC |

O ~1 GB por pessoa é a parte que se sente. Se incomodar, o candidato óbvio é
`data/pdfs` (o arquivo de licenças): pode ficar numa biblioteca à parte,
sincronizada só por quem precisa, e o pipeline aponta para lá com
`EPAL_TUA_PROCESSED` / `EPAL_LURH_PROCESSED`.

---

## 5. Duas coisas que só passam a ser perigosas agora

Enquanto a pasta era o OneDrive de UMA pessoa, isto era teórico. Numa
biblioteca de equipa deixa de ser.

**Duas pessoas a processar ao mesmo tempo.** O `PROCESSAR_AGORA.bat` existe para ser
clicado por operadores, e nada impedia dois de o clicarem no mesmo minuto: os
dois liam o master, os dois reescreviam-no inteiro, e o último a gravar
apagava as licenças do outro — sem erro nenhum. Só o vigia automático estava
protegido. Agora as execuções manuais partilham o mesmo cadeado
(`data/_staging/run.lock`): a segunda pessoa vê

```
Ja esta a correr no computador de <nome> (ha 3 min). NAO corri, para nao
estragar o master. Espere que termine e tente outra vez.
```

e sai com o código 2, sem tocar em nada. Não é um cadeado distribuído a
sério — o SharePoint não dá nenhuma primitiva atómica entre máquinas — mas
cobre o caso real, que é gente a trabalhar com minutos de diferença.

**Os ficheiros de dados também são marcadores na nuvem.** Fixar só o
`runtime` não chegava: os master, o feed do Power BI e o código são lidos e
escritos em todas as execuções. O `Automation files\CONFIGURAR_ESTE_PC.bat` passa a fixar
`runtime`, `src`, `data\powerbi` e `reports`. Fica de fora `data\pdfs` — são
~600 MB de arquivo que só se escreve, e não vale a pena obrigar cada portátil
a ter tudo isso em disco.

---

## 6. Coisas que já foram verificadas

- **Nomes de ficheiros**: nenhum tem caracteres proibidos pelo SharePoint
  (`" * : < > ? / \ |`), espaços ou pontos no fim, nem nomes reservados.
  Verificado em 11 179 nomes.
- **Número de ficheiros**: ~11 200, muito abaixo do limite de sincronização.
- **O Python é mesmo portátil**: o projeto foi copiado para `C:\Temp` e uma
  extração LURH completa correu de lá — o `.bat` encontrou o interpretador
  no sítio novo e a licença foi arquivada no caminho novo. A mudança de
  pasta, por si só, não parte nada.
