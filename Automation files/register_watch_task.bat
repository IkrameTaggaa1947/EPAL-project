@echo off
REM ===================================================================
REM  UM vigia automatico para OS DOIS extratores (TUA + LURH).
REM  Duplo-clique UMA VEZ. Regista uma tarefa do Windows que verifica as
REM  duas pastas a cada 1 minuto, em segundo plano, sem janela, e
REM  continua a funcionar depois de reiniciar o PC.
REM
REM   - PDFs TUA  em  ..\..\Drop_New_Licenses\new_TUA_licences
REM   - PDFs LURH em  ..\..\Drop_New_Licenses\new_LURH_licences
REM
REM  Para desligar, duplo-clique em unregister_watch_task.bat.
REM
REM  NOTA: a tarefa fica registada SO NESTE PC e aponta para a copia da
REM  pasta sincronizada NESTE PC. Cada pessoa que queira o vigia
REM  automatico tem de correr este ficheiro no seu proprio computador.
REM ===================================================================
setlocal

call "%~dp0_find_python.bat"
if not defined EPAL_PYW (
    call "%~dp0_no_python.bat"
    exit /b 1
)

REM Resolver para caminho absoluto: a tarefa agendada corre sem pasta atual,
REM e o resolvedor devolve um caminho com ".." que o schtasks nao gosta.
REM Se for um comando simples ("pythonw"), fica como esta.
set "PYW=%EPAL_PYW%"
if exist "%EPAL_PYW%" for %%I in ("%EPAL_PYW%") do set "PYW=%%~fI"
set "SCRIPT=%~dp0..\src\automation_of_extraction\watch_all.py"
REM  o schtasks nao aceita ".." no caminho: resolver para absoluto
for %%I in ("%SCRIPT%") do set "SCRIPT=%%~fI"

REM Remover tarefas antigas por dominio, para nao correrem tres vigias.
schtasks /Delete /TN "EPAL TUA Watch"  /F >nul 2>&1
schtasks /Delete /TN "EPAL LURH Watch" /F >nul 2>&1

schtasks /Create /TN "EPAL Watch" /TR "\"%PYW%\" \"%SCRIPT%\"" /SC MINUTE /MO 1 /F
set "RC=%errorlevel%"

REM A tarefa corre a cada minuto, mas um lote pode demorar muito mais.
REM Sem isto, a politica de instancias multiplas fica no valor por omissao
REM do schtasks - e duas instancias a escrever no mesmo master corrompem-no.
REM O schtasks nao tem flag para isto; so se define pelo PowerShell.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "try { $t = Get-ScheduledTask -TaskName 'EPAL Watch' -ErrorAction Stop; $t.Settings.MultipleInstances = 'IgnoreNew'; Set-ScheduledTask -InputObject $t | Out-Null; Write-Host '        instancias multiplas: IgnoreNew (nao arranca uma segunda copia)' } catch { Write-Host '        AVISO: nao foi possivel fixar a politica de instancias multiplas.' }"

if "%RC%"=="0" (
    echo.
    echo SUCESSO. As duas pastas passam a ser processadas automaticamente a
    echo cada 1 minuto, em segundo plano, sem janela - e depois de cada reinicio.
    echo Para desligar mais tarde, corra  unregister_watch_task.bat
) else (
    echo.
    echo Nao foi possivel criar a tarefa.
    echo Clique com o botao direito neste ficheiro e escolha
    echo "Executar como administrador".
)
echo.
pause
endlocal
