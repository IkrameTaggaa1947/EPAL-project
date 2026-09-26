@echo off
setlocal
REM ===================================================================
REM  Comparar os criterios de uma licenca ANTIGA com a NOVA (mesma ETAR).
REM  Duplo-clique para correr.
REM
REM  Funciona SEMPRE, mesmo sem o AMALIA instalado: nesse caso a camada
REM  deterministica corre na mesma e as colunas do AMALIA ficam a
REM  "PENDENTE" no Excel.
REM
REM  Para ligar o AMALIA neste PC, edite:
REM     %LOCALAPPDATA%\EPAL\epal.local.ini    seccao [amalia]
REM  Nao ha nada para configurar dentro da pasta partilhada.
REM ===================================================================
cd /d "%~dp0..\src\criteria_LLM_automation"
call "%~dp0_find_python.bat"
if not defined EPAL_PY (
    call "%~dp0_no_python.bat"
    exit /b 1
)

"%EPAL_PY%" compare_legislacao.py %*

echo.
echo  O Excel foi gravado em  outputs\comparacoes\
echo.
pause
endlocal
