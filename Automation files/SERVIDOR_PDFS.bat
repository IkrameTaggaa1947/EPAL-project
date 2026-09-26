@echo off
REM ===================================================================
REM  EPAL - servidor local dos PDFs
REM
REM  O Power BI NAO abre ficheiros locais a partir de um botao. Este
REM  servidor da aos PDFs um endereco http:// nesta maquina, e o botao
REM  "Abrir licenca" passa a funcionar.
REM
REM  Duplo-clique e DEIXE ESTA JANELA ABERTA enquanto usa o dashboard.
REM  Depois da mudanca para o SharePoint isto deixa de ser preciso.
REM ===================================================================
cd /d "%~dp0..\"
call "%~dp0_find_python.bat"
if not defined EPAL_PY (
    call "%~dp0_no_python.bat"
    exit /b 1
)
"%EPAL_PY%" "%~dp0..\src\Common_Code_PowerBI\servidor_pdfs.py"
pause
