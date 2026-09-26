@echo off
REM Mensagem unica para quando nao ha Python disponivel.
echo.
echo  ============================================================
echo   NAO FOI ENCONTRADO NENHUM PYTHON
echo  ============================================================
echo.
echo   Normalmente isto NAO devia acontecer: o projeto traz um
echo   Python portatil em  EPAL-project\runtime\python .
echo.
echo   Causas provaveis:
echo     - a pasta ainda nao acabou de sincronizar (espere e tente
echo       outra vez);
echo     - o Python portatil ainda nao foi criado. Nesse caso,
echo       alguem com internet deve fazer duplo-clique UMA VEZ em:
echo.
echo           EPAL-project\Automation files\INSTALAR_PYTHON_PORTATIL.bat
echo.
pause
goto :eof
