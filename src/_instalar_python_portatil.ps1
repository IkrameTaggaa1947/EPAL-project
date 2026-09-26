# ===================================================================
#  Cria (ou repara) o Python PORTATIL dentro da pasta do projeto.
#
#  PORQUE: os colegas nao podem/nao querem instalar Python nos
#  portateis. Este Python vive dentro da pasta do projeto, sincroniza
#  com ela, e nao precisa de instalacao nem de direitos de admin.
#
#  Corre UMA VEZ, num computador com internet. Depois a pasta
#  sincroniza para toda a gente e os .bat encontram-no sozinhos.
#
#  Se ja existir, oferece REPARAR (reinstala as bibliotecas em falta
#  sem voltar a descarregar o Python).
# ===================================================================
$ErrorActionPreference = "Stop"

$Versao  = "3.13.1"                     # tem de ser 3.13.x (bibliotecas cp313)

# Checksum publicado pela python.org para python-$Versao-embed-amd64.zip, em
#   https://www.python.org/downloads/release/python-3131/
# AO MUDAR $Versao, MUDAR TAMBEM ISTO (a pagina da nova versao tem o valor).
# O TLS ja garante que estamos a falar com a python.org; isto apanha um
# download corrompido a meio e um ficheiro trocado no caminho.
$ZipMD5  = "d5c8030976b5eaf55ed6b321c073dda7"
$Raiz    = Split-Path -Parent $PSScriptRoot           # ...\EPAL-project
$Destino = Join-Path $Raiz "runtime\python"
$Req     = Join-Path $Raiz "requirements.txt"
$Temp    = Join-Path $env:TEMP "epal_python_portatil"
$PyExe   = Join-Path $Destino "python.exe"

$UrlZip  = "https://www.python.org/ftp/python/$Versao/python-$Versao-embed-amd64.zip"
$UrlPip  = "https://bootstrap.pypa.io/get-pip.py"

function Instalar-Bibliotecas {
    # SEM --quiet: se algo falhar, tem de se ver. Com repeticoes, porque uma
    # falha de rede a meio deixa o ambiente incompleto (foi o que aconteceu
    # na primeira instalacao: faltaram dateutil/six e o pandas nao importava).
    Write-Host "  A instalar as bibliotecas (cerca de 43 MB)..."
    & $PyExe -m pip install --no-warn-script-location --retries 5 --timeout 60 -r $Req
    if ($LASTEXITCODE -ne 0) { throw "pip install falhou." }

    Write-Host ""
    Write-Host "  A confirmar que nao falta nenhuma dependencia..."
    & $PyExe -m pip check
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Faltavam dependencias - a tentar corrigir..." -ForegroundColor Yellow
        & $PyExe -m pip install --no-warn-script-location --retries 5 --timeout 60 --upgrade -r $Req
        & $PyExe -m pip check
        if ($LASTEXITCODE -ne 0) { throw "Continuam a faltar dependencias (ver lista acima)." }
    }

    Write-Host ""
    Write-Host "  A verificar os modulos um a um..."
    & $PyExe -c @"
import importlib, sys
faltam = []
for m in ('fitz', 'pymupdf', 'pandas', 'numpy', 'openpyxl', 'dateutil'):
    try:
        importlib.import_module(m)
        print('        OK    ' + m)
    except Exception as e:
        faltam.append(m); print('        FALHA ' + m + '  ' + str(e))
sys.exit(1 if faltam else 0)
"@
    if ($LASTEXITCODE -ne 0) { throw "Ha modulos que nao importam - a instalacao nao ficou boa." }
}

Write-Host ""
Write-Host " ===============================================================" -ForegroundColor Cyan
Write-Host "   EPAL - Python portatil"                                       -ForegroundColor Cyan
Write-Host " ===============================================================" -ForegroundColor Cyan
Write-Host "   Destino: $Destino"
Write-Host ""

# ---------- Ja existe: reparar ou refazer --------------------------
if (Test-Path $PyExe) {
    Write-Host "  Ja existe um Python portatil nesta pasta." -ForegroundColor Yellow
    Write-Host "    [R] Reparar  - reinstala so as bibliotecas (rapido, recomendado)"
    Write-Host "    [C] Criar de novo - apaga tudo e descarrega outra vez"
    Write-Host "    [N] Nao fazer nada"
    $r = Read-Host "  Escolha (R/C/N)"
    if ($r -match '^[Rr]') { Instalar-Bibliotecas; Write-Host ""; Write-Host "  REPARADO." -ForegroundColor Green; return }
    elseif ($r -match '^[Cc]') { Remove-Item -Recurse -Force $Destino }
    else { Write-Host "  Cancelado."; return }
}

# TLS 1.2 - alguns Windows nao o usam por omissao e a descarga falha.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
New-Item -ItemType Directory -Force -Path $Temp, $Destino | Out-Null

Write-Host "  [1/5] A descarregar o Python $Versao (cerca de 11 MB)..."
$Zip = Join-Path $Temp "python-embed.zip"
Invoke-WebRequest -Uri $UrlZip -OutFile $Zip -UseBasicParsing
$obtido = (Get-FileHash -Path $Zip -Algorithm MD5).Hash
if ($obtido -ne $ZipMD5) {
    Remove-Item -Force $Zip -ErrorAction SilentlyContinue
    throw ("O ficheiro descarregado nao corresponde ao publicado pela python.org.`n" +
           "  esperado: $ZipMD5`n" +
           "  obtido  : $obtido`n" +
           "Tente outra vez. Se persistir, confirme o valor em " +
           "https://www.python.org/downloads/release/python-$($Versao -replace '\.','')/")
}
Write-Host "        descarregado e verificado."

Write-Host "  [2/5] A extrair..."
Expand-Archive -Path $Zip -DestinationPath $Destino -Force
Write-Host "        extraido para runtime\python"

# O Python embutivel traz um ._pth que RESTRINGE o sys.path e desliga o
# site-packages. Sem este passo o pip instala mas nada importa.
Write-Host "  [3/5] A configurar o caminho das bibliotecas (._pth)..."
$Pth = Get-ChildItem -Path $Destino -Filter "python*._pth" | Select-Object -First 1
if (-not $Pth) { throw "Nao encontrei o ficheiro ._pth no Python embutivel." }
$linhas = Get-Content $Pth.FullName | ForEach-Object {
    if ($_ -match '^\s*#\s*import\s+site\s*$') { "import site" } else { $_ }
}
if ($linhas -notcontains "Lib\site-packages") { $linhas += "Lib\site-packages" }
if ($linhas -notcontains "import site")       { $linhas += "import site" }
Set-Content -Path $Pth.FullName -Value $linhas -Encoding ASCII
Write-Host "        $($Pth.Name) configurado."

Write-Host "  [4/5] A instalar o pip..."
$GetPip = Join-Path $Temp "get-pip.py"
Invoke-WebRequest -Uri $UrlPip -OutFile $GetPip -UseBasicParsing
& $PyExe $GetPip --no-warn-script-location --quiet
if ($LASTEXITCODE -ne 0) { throw "Falhou a instalacao do pip." }
Write-Host "        pip instalado."

Write-Host "  [5/5]"
Instalar-Bibliotecas

Remove-Item -Recurse -Force $Temp -ErrorAction SilentlyContinue

Write-Host ""
Write-Host " ===============================================================" -ForegroundColor Green
Write-Host "   PRONTO - Python portatil criado e verificado"                  -ForegroundColor Green
Write-Host " ===============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   A partir de agora NENHUM computador precisa de instalar Python."
Write-Host "   Assim que a pasta sincronizar, os .bat encontram este Python"
Write-Host "   sozinhos e tudo funciona por duplo-clique."
Write-Host ""
