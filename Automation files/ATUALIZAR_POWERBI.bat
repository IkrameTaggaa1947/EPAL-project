@echo off
setlocal
REM ===================================================================
REM  Aponta o dashboard Power BI para a pasta de dados DESTE PC.
REM
REM  Corra isto se o Power BI der erro do genero:
REM     "nao foi possivel encontrar a origem de dados"
REM     "a coluna X nao foi encontrada"
REM
REM  O Power BI Desktop tem de estar FECHADO.
REM ===================================================================
cd /d "%~dp0..\"
call "%~dp0_find_python.bat"
if not defined EPAL_PY (
    call "%~dp0_no_python.bat"
    exit /b 1
)

echo.
echo  A verificar o caminho dos dados para este PC...
echo.
"%EPAL_PY%" "%~dp0..\src\Common_Code_PowerBI\fix_powerbi_datafolder.py"
echo.
echo  Pode agora abrir o dashboard e clicar em Atualizar (Refresh).
echo.
pause
endlocal
