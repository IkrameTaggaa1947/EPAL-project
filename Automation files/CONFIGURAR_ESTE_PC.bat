@echo off
setlocal
REM ===================================================================
REM   EPAL - CONFIGURAR ESTE COMPUTADOR
REM   Duplo-clique. Correr UMA VEZ por computador.
REM
REM   NAO instala nada e NAO precisa de direitos de administrador.
REM   O Python vem dentro da pasta do projeto (runtime\python).
REM
REM   O que faz:
REM     1. Confirma que encontra o Python do projeto.
REM     2. Fixa o Python neste disco (o OneDrive deixa de o esvaziar).
REM     3. Confirma que TODAS as bibliotecas importam mesmo.
REM     4. Cria o ficheiro de definicoes SO deste PC.
REM     5. Aponta o dashboard Power BI para a pasta deste PC.
REM     e mostra um diagnostico final.
REM ===================================================================
cd /d "%~dp0..\"
REM  Caminho absoluto da raiz, sem o "\.." no meio - usado no passo 2 para
REM  saber se a pasta esta ou nao dentro do OneDrive.
set "RAIZ=%CD%"

echo.
echo  ===============================================================
echo    EPAL - configuracao deste computador
echo  ===============================================================
echo.
echo    Pasta do projeto : %~dp0..\
echo    Definicoes locais: %LOCALAPPDATA%\EPAL
echo.

REM ---------- 1. Python ----------------------------------------------
echo  [1/5] A procurar o Python do projeto...
call "%~dp0_find_python.bat"
if not defined EPAL_PY (
    call "%~dp0_no_python.bat"
    exit /b 1
)
if exist "%~dp0..\runtime\python\python.exe" (
    echo        encontrado o Python PORTATIL do projeto - nada a instalar.
) else (
    echo        AVISO: o Python portatil ainda nao existe nesta pasta.
    echo        A usar o Python deste computador. Para deixar de depender
    echo        disso, alguem com internet deve correr UMA VEZ:
    echo            INSTALAR_PYTHON_PORTATIL.bat
)
"%EPAL_PY%" --version
echo.

REM ---------- 2. Manter o Python sempre neste disco --------------------
REM  O OneDrive guarda os ficheiros "a pedido": numa maquina nova os 8000
REM  ficheiros do runtime sao so marcadores na nuvem, e o OneDrive pode
REM  voltar a esvazia-los mais tarde para poupar espaco - partindo um
REM  Python que ja funcionava. "attrib +P" e o mesmo que a opcao
REM  "Manter sempre neste dispositivo" do menu do OneDrive. Nao precisa
REM  de direitos de administrador.
echo  [2/5] A fixar o Python neste disco (OneDrive: manter sempre)...
REM  Sem blocos "if" encaixados: com parenteses aninhados o cmd.exe avalia
REM  o errorlevel no sitio errado. Guardar em %RC% e comparar e sempre seguro.
REM  Fixa-se tambem o que o pipeline LE ou ESCREVE em cada execucao: os
REM  ficheiros master, o feed do Power BI, o codigo e o dashboard. Fica de
REM  fora data\pdfs - sao ~600 MB de arquivo que so se escreve, e nao vale a
REM  pena obrigar cada portatil a ter tudo isso em disco.
REM
REM  SO FAZ SENTIDO DENTRO DO ONEDRIVE. O "attrib +P" e uma marca do OneDrive
REM  ("manter sempre neste dispositivo"); numa pasta de rede (W:\, \\servidor)
REM  nao quer dizer nada -- e percorrer 8000 ficheiros por SMB fica pendurado
REM  durante imenso tempo, que foi exactamente o que aconteceu num portatil
REM  com o projeto em W:\DOS\...  Fora do OneDrive nao ha desidratacao de
REM  ficheiros, portanto nao ha nada a fixar: passa-se a frente.
echo "%RAIZ%" | find /i "OneDrive" >nul
if errorlevel 1 goto :fora_onedrive
if not exist "%~dp0..\runtime\python\python.exe" goto :sem_portatil
attrib +P /S /D "%~dp0..\runtime\*"      >nul 2>&1
set "RC=%errorlevel%"
attrib +P /S /D "%~dp0..\src\*"          >nul 2>&1
attrib +P /S /D "%~dp0..\data\powerbi\*" >nul 2>&1
attrib +P /S /D "%~dp0..\reports\*"      >nul 2>&1
if "%RC%"=="0" (
    echo        fixado - o OneDrive deixa de o esvaziar.
    echo        ^(runtime, src, data\powerbi e reports; data\pdfs fica a pedido^)
) else (
    echo        AVISO: nao consegui fixar. Se o Python falhar mais tarde,
    echo        clique com o botao direito na pasta runtime e escolha
    echo        "Manter sempre neste dispositivo".
)
goto :fim_fixar
:fora_onedrive
echo        a pasta nao esta no OneDrive - nada a fixar.
echo        ^(so o OneDrive esvazia ficheiros; numa pasta de rede ou local
echo         eles estao sempre ca^)
goto :fim_fixar
:sem_portatil
echo        (sem Python portatil nesta pasta - nada a fixar)
:fim_fixar
echo.

REM ---------- 3. Confirmar que as bibliotecas estao todas --------------
REM  "python --version" funciona mesmo com a pasta a meio da sincronizacao.
REM  O que parte a extracao e faltar UMA biblioteca, e isso so se ve
REM  tentando importa-las - que e o que este passo faz.
echo  [3/5] A confirmar que as bibliotecas estao todas...
"%EPAL_PY%" "%~dp0..\src\_verificar_bibliotecas.py"
set "RC=%errorlevel%"
if not "%RC%"=="0" (
    echo.
    echo        A instalacao do Python nao esta completa neste computador.
    echo        Normalmente e a pasta ainda a sincronizar: espere e repita.
    echo        Se continuar, alguem com internet deve correr UMA VEZ:
    echo            INSTALAR_PYTHON_PORTATIL.bat   ^(opcao R - Reparar^)
    echo.
    pause
    exit /b 1
)
echo.

REM ---------- 4. Definicoes locais ------------------------------------
echo  [4/5] A criar o ficheiro de definicoes deste PC...
"%EPAL_PY%" -c "import sys; sys.path.insert(0, r'%~dp0..\src'); import epal_config as c; print('       ', c.ensure_local_config())"
echo.

REM ---------- 5. Power BI ---------------------------------------------
echo  [5/5] A apontar o dashboard Power BI para este PC...
echo        (cria a ligacao estavel C:\EPAL\powerbi, igual em todos os PCs,
echo         para o dashboard deixar de se partir quando a pasta sincroniza)
echo        (o Power BI Desktop deve estar FECHADO)
"%EPAL_PY%" "%~dp0..\src\Common_Code_PowerBI\fix_powerbi_datafolder.py"
echo.

REM ---------- Diagnostico ----------------------------------------------
echo  Diagnostico final:
echo.
"%EPAL_PY%" "%~dp0..\src\epal_config.py"

echo.
echo  ===============================================================
echo    CONFIGURACAO TERMINADA
echo  ===============================================================
echo.
echo    Ja pode:
echo      - por PDFs em  Drop_New_Licenses\new_TUA_licences  (ou LURH)
echo        e fazer duplo-clique em  src\Extraction_Code_TUA\PROCESSAR_AGORA.bat
echo      - abrir o dashboard em  reports\EPAL_Licencas_Dashboard
echo        e clicar em Atualizar (Refresh)
echo.
echo    ATENCAO - vigia automatico:
echo      NAO registe o vigia automatico (register_watch_task.bat) em
echo      varios computadores ao mesmo tempo. A pasta e partilhada: se
echo      dois PCs processarem os mesmos PDFs em simultaneo, os ficheiros
echo      master podem ficar corrompidos. Escolham UM computador para
echo      isso. Nos restantes, use PROCESSAR_AGORA.bat quando
echo      precisar.
echo.
pause
endlocal
