# Build tools live outside OneDrive; end users need only the published files.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$toolsRoot = Join-Path $env:LOCALAPPDATA 'PainelGAQ/build'
$nodeCmd = Get-Command node.exe -ErrorAction SilentlyContinue
$node = if ($nodeCmd) { $nodeCmd.Source } else { Join-Path $toolsRoot 'node/node.exe' }
if (-not (Test-Path -LiteralPath $node)) { throw 'Node 22.19+ necessario apenas na maquina de compilacao. Instale ou coloque a versao portatil em LOCALAPPDATA/PainelGAQ/build/node.' }
$env:PATH = (Split-Path -Parent $node) + ';' + $env:PATH
$env:GAQ_BUILD_MODULES = Join-Path $toolsRoot 'node_modules'
New-Item -ItemType Directory -Path $toolsRoot -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $root 'package.json') -Destination $toolsRoot -Force
Copy-Item -LiteralPath (Join-Path $root 'package-lock.json') -Destination $toolsRoot -Force
Push-Location $root
try {
    & npm.cmd ci --prefix $toolsRoot --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao preparar ferramentas de compilacao.' }
    & $node --use-system-ca scripts/build.mjs
    if ($LASTEXITCODE -ne 0) { throw 'Falha na compilacao.' }
} finally { Pop-Location }
