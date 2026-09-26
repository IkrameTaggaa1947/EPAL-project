@echo off
REM ===================================================================
REM  Cria um Python PORTATIL dentro da pasta do projeto.
REM
REM  CORRER UMA VEZ, num computador com internet.
REM  Depois disso, mais nenhum computador precisa de instalar Python:
REM  a pasta sincroniza e os outros .bat encontram-no sozinhos.
REM
REM  Demora 2 a 4 minutos e ocupa cerca de 150 MB na pasta partilhada.
REM ===================================================================
cd /d "%~dp0..\"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\src\_instalar_python_portatil.ps1"
echo.
pause
