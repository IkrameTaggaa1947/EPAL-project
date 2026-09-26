@echo off
REM ===================================================================
REM  EPAL - DIAGNOSTICO deste computador.
REM
REM  Duplo-clique e envie TUDO o que aparece. Funciona mesmo quando nada
REM  mais funciona: nao precisa do Python nem de configuracao nenhuma.
REM
REM  Escrito sem blocos ^( ^) encaixados de proposito - o cmd.exe avalia
REM  as variaveis no sitio errado dentro deles.
REM ===================================================================
setlocal enabledelayedexpansion
cd /d "%~dp0.."

echo ===============================================================
echo   EPAL - DIAGNOSTICO
echo ===============================================================
echo.
echo  Computador : %COMPUTERNAME%
echo  Utilizador : %USERNAME%
echo  Data       : %DATE% %TIME%
echo.

REM ---------- 1. Onde esta o projeto, e o caminho e curto que baste? ----
set "RAIZ=%CD%"
echo  [1] PASTA DO PROJETO
echo      %RAIZ%
call :comprimento "%RAIZ%"
echo      comprimento: !TAM! caracteres  ^(o limite util e 113^)
if !TAM! GTR 113 echo      *** AVISO: caminho demasiado longo - havera ficheiros em falta.
echo.

REM ---------- 2. As pastas de topo estao todas la? ---------------------
echo  [2] PASTAS DE TOPO
for %%D in (src data reports runtime outputs Drop_New_Licenses "Automation files") do (
    if exist "%RAIZ%\%%~D\" (echo      OK       %%~D) else (echo      EM FALTA %%~D)
)
echo.

REM ---------- 3. O Python portatil ------------------------------------
echo  [3] PYTHON DO PROJETO
if exist "%RAIZ%\runtime\python\python.exe" goto :py_ok
echo      EM FALTA  runtime\python\python.exe
echo      *** E por isto que os .bat dizem "NAO FOI ENCONTRADO NENHUM PYTHON".
echo      *** Causa habitual: a pasta ainda nao acabou de sincronizar.
set "EPAL_PY="
goto :py_fim
:py_ok
echo      OK        runtime\python\python.exe
set "EPAL_PY=%RAIZ%\runtime\python\python.exe"
"%EPAL_PY%" --version
:py_fim
echo.

REM ---------- 4. A ligacao que o Power BI usa --------------------------
echo  [4] LIGACAO DO POWER BI  ^(C:\EPAL\powerbi^)
if exist "C:\EPAL\powerbi\" goto :jun_ok
echo      NAO EXISTE
echo      *** O Power BI nao vai encontrar os ficheiros Excel.
echo      *** Solucao: fechar o Power BI e correr CONFIGURAR_ESTE_PC.bat
goto :jun_fim
:jun_ok
dir "C:\EPAL" 2>nul | findstr /i "JUNCTION"
if exist "C:\EPAL\powerbi\Licenses.xlsx" (echo      OK  Licenses.xlsx visivel atraves da ligacao) else (echo      *** A ligacao existe mas esta PARTIDA - aponta para uma pasta que ja nao existe.)
:jun_fim
echo.

REM ---------- 5. Consigo escrever? ------------------------------------
echo  [5] PERMISSAO DE ESCRITA
break > "%RAIZ%\_teste_escrita.tmp" 2>nul
if exist "%RAIZ%\_teste_escrita.tmp" goto :w_ok
echo      NAO  - a pasta esta apenas para LEITURA neste computador.
echo      *** Nenhuma alteracao ao codigo resolve isto: quem partilhou a
echo      *** pasta tem de dar permissao de EDICAO a este utilizador.
goto :w_fim
:w_ok
del "%RAIZ%\_teste_escrita.tmp" >nul 2>&1
echo      OK   consigo escrever na pasta do projeto.
:w_fim
echo.

REM ---------- 6. Definicoes locais deste PC ----------------------------
echo  [6] DEFINICOES LOCAIS
if exist "%LOCALAPPDATA%\EPAL\epal.local.ini" (echo      OK       %LOCALAPPDATA%\EPAL\epal.local.ini) else (echo      EM FALTA - este PC nunca correu CONFIGURAR_ESTE_PC.bat)
echo.

REM ---------- 7. Os mapas conseguem chegar aos servidores? -------------
REM  Um mapa TOTALMENTE BRANCO quer quase sempre dizer que o visual abriu
REM  mas os "tiles" (as imagens do mapa) nao vieram. Isso e rede, nao e
REM  o Power BI. O Azure Maps usa atlas.microsoft.com; o mapa normal usa
REM  o Bing. Testar os dois separa as duas coisas.
echo  [7] ACESSO AOS SERVIDORES DOS MAPAS
if not defined EPAL_PY goto :net_fim
"%EPAL_PY%" "%RAIZ%\src\_verificar_mapas.py"
goto :net_ok
:net_fim
echo      ^(sem Python - nao foi possivel testar^)
:net_ok
echo.

REM ---------- 8. Verificacao completa, se houver Python ----------------
echo  [8] VERIFICACAO COMPLETA
if not defined EPAL_PY goto :fim
"%EPAL_PY%" "%RAIZ%\src\_verificar_migracao.py"

:fim
echo.
echo ===============================================================
echo   FIM - copie TUDO o que esta acima e envie.
echo ===============================================================
echo.
pause
endlocal
goto :eof

REM ---------- subrotina: comprimento de uma string ---------------------
:comprimento
set "S=%~1"
set TAM=0
:cl
if defined S set "S=!S:~1!" & set /a TAM+=1 & goto :cl
goto :eof
