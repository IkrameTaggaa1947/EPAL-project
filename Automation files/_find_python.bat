@echo off
REM ===================================================================
REM  Descobre que Python usar neste PC e devolve-o em %EPAL_PY%
REM  (e a versao sem janela em %EPAL_PYW%).
REM  Chamado pelos outros .bat com:  call "%~dp0_find_python.bat"
REM
REM  Ordem de preferencia:
REM    1. Python PORTATIL da pasta partilhada  (runtime\python)
REM       -> e o normal: ninguem precisa de instalar nada.
REM    2. Ambiente virtual local deste PC      (%LOCALAPPDATA%\EPAL\venv)
REM    3. Python instalado no PC
REM    4. Lancador "py"
REM
REM  NOTA: o .venv que possa existir dentro da pasta partilhada e
REM  IGNORADO de proposito - foi criado noutro computador e nao
REM  funciona neste.
REM ===================================================================
set "EPAL_PY="
set "EPAL_PYW="

REM --- 1. Python portatil (dentro do projeto, sincronizado) ---------
if exist "%~dp0..\runtime\python\python.exe" (
    set "EPAL_PY=%~dp0..\runtime\python\python.exe"
    set "EPAL_PYW=%~dp0..\runtime\python\pythonw.exe"
    goto :eof
)

REM --- 2. Ambiente virtual local a este PC --------------------------
if exist "%LOCALAPPDATA%\EPAL\venv\Scripts\python.exe" (
    set "EPAL_PY=%LOCALAPPDATA%\EPAL\venv\Scripts\python.exe"
    set "EPAL_PYW=%LOCALAPPDATA%\EPAL\venv\Scripts\pythonw.exe"
    goto :eof
)

REM --- 3. Python instalado ------------------------------------------
where python >nul 2>&1
if %errorlevel%==0 (
    set "EPAL_PY=python"
    set "EPAL_PYW=pythonw"
    goto :eof
)

REM --- 4. Lancador py -----------------------------------------------
where py >nul 2>&1
if %errorlevel%==0 (
    set "EPAL_PY=py"
    set "EPAL_PYW=pyw"
)
goto :eof
