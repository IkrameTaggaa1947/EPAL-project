@echo off
REM ===================================================================
REM  OS DOIS extratores - MODO VIGIA (fica a correr, janela visivel).
REM  Duplo-clique uma vez e deixe a janela aberta: qualquer PDF TUA ou
REM  LURH que ponha na respetiva pasta e processado automaticamente.
REM  Feche a janela para parar.
REM  (Para correr em segundo plano, use register_watch_task.bat.)
REM ===================================================================
cd /d "%~dp0..\src\automation_of_extraction"
call "%~dp0_find_python.bat"
if not defined EPAL_PY (
    call "%~dp0_no_python.bat"
    exit /b 1
)

"%EPAL_PY%" watch_all.py --watch

pause
