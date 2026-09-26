@echo off
REM ===================================================================
REM  EPAL - processar AGORA os PDFs que estao a espera (TUA + LURH).
REM  Duplo-clique: corre UMA vez e sai.
REM
REM  1) Ponha os PDFs em  ..\Drop_New_Licenses\new_TUA_licences
REM                 ou    ..\Drop_New_Licenses\new_LURH_licences
REM  2) Faca duplo-clique neste ficheiro.
REM  3) No fim: as licencas boas vao para ..\data\pdfs\NEW\
REM     e so as que falharam ficam em data\review.
REM
REM  SUBSTITUI os dois ficheiros que havia por regime. O watch_all.py sem
REM  argumentos ja percorre AS DUAS pastas uma vez e sai; uma pasta vazia
REM  demora dois segundos, por isso nao havia nada a ganhar em ter um
REM  ficheiro para cada regime.
REM
REM  Para deixar a vigiar com janela aberta : watch_all.bat
REM  Para vigiar em segundo plano, sempre   : register_watch_task.bat
REM ===================================================================
cd /d "%~dp0..\src\automation_of_extraction"
call "%~dp0_find_python.bat"
if not defined EPAL_PY (
    call "%~dp0_no_python.bat"
    exit /b 1
)

"%EPAL_PY%" watch_all.py

echo.
echo Terminado. Veja as mensagens acima.
pause
