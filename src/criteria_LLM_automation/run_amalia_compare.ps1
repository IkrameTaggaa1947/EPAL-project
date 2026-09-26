# ===================================================================
#  OBSOLETO - nao usar.
#
#  Este ficheiro tinha o endereco e a chave do AMALIA escritos a mao
#  ("localhost:8000"), o que so funcionava no computador onde foi
#  criado. As definicoes do AMALIA passaram para um ficheiro proprio
#  de cada maquina:
#
#      %LOCALAPPDATA%\EPAL\epal.local.ini     ->  seccao [amalia]
#
#  Para correr a comparacao, use:
#
#      COMPARAR_CRITERIOS.bat        (duplo-clique)
#
#  Mantido apenas para nao partir atalhos antigos.
# ===================================================================
Write-Host ""
Write-Host "  Este script foi substituido por COMPARAR_CRITERIOS.bat" -ForegroundColor Yellow
Write-Host "  As definicoes do AMALIA estao agora em:" -ForegroundColor Yellow
Write-Host "      $env:LOCALAPPDATA\EPAL\epal.local.ini" -ForegroundColor Yellow
Write-Host ""
$bat = Join-Path $PSScriptRoot "COMPARAR_CRITERIOS.bat"
if (Test-Path $bat) {
    $ans = Read-Host "  Quer correr COMPARAR_CRITERIOS.bat agora? (S/N)"
    if ($ans -match '^[SsYy]') { & $bat @args }
}
