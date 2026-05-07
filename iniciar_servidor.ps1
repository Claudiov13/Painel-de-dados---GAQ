$ErrorActionPreference = "Stop"

$Projeto = Split-Path -Parent $MyInvocation.MyCommand.Path
$NodeCandidates = @(
    (Join-Path $Projeto "node-portable\node.exe"),
    (Join-Path $Projeto "nodejs\node.exe"),
    (Join-Path $Projeto "tools\node\node.exe"),
    "$env:ProgramFiles\nodejs\node.exe",
    "$env:LOCALAPPDATA\Programs\nodejs\node.exe",
    "$env:ProgramFiles(x86)\nodejs\node.exe"
)

$NodeCommand = Get-Command node -ErrorAction SilentlyContinue
if ($NodeCommand) {
    $NodePath = $NodeCommand.Source
} else {
    $NodePath = $NodeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $NodePath) {
        $NodePath = Get-ChildItem -LiteralPath $Projeto -Recurse -Filter node.exe -ErrorAction SilentlyContinue |
            Where-Object {
                $_.FullName -like (Join-Path $Projeto "node-portable\*") -or
                $_.FullName -like (Join-Path $Projeto "nodejs\*") -or
                $_.FullName -like (Join-Path $Projeto "tools\node\*")
            } |
            Select-Object -First 1 -ExpandProperty FullName
    }
}

if (-not $NodePath) {
    throw "Node.js nao encontrado. Instale o Node.js LTS antes de iniciar o servidor."
}

$LogDir = Join-Path $Projeto "logs"
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

$OutLog = Join-Path $LogDir "server-out.log"
$ErrLog = Join-Path $LogDir "server-err.log"
$ServerScript = Join-Path $Projeto "server.js"

while ($true) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $OutLog -Value "[$timestamp] Iniciando servidor..."

    $process = Start-Process `
        -FilePath $NodePath `
        -ArgumentList "`"$ServerScript`"" `
        -WorkingDirectory $Projeto `
        -RedirectStandardOutput $OutLog `
        -RedirectStandardError $ErrLog `
        -Wait `
        -PassThru `
        -NoNewWindow

    $exitCode = $process.ExitCode
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $ErrLog -Value "[$timestamp] Servidor encerrado com codigo $exitCode. Reiniciando em 5 segundos..."
    Start-Sleep -Seconds 5
}
