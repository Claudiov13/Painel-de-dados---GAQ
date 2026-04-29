$ErrorActionPreference = "Stop"

$Projeto = Split-Path -Parent $MyInvocation.MyCommand.Path
$Launcher = Join-Path $Projeto "iniciar_servidor.ps1"
$NodeCandidates = @(
    (Join-Path $Projeto "node-portable\node.exe"),
    (Join-Path $Projeto "nodejs\node.exe"),
    (Join-Path $Projeto "tools\node\node.exe"),
    "$env:ProgramFiles\nodejs\node.exe",
    "$env:LOCALAPPDATA\Programs\nodejs\node.exe",
    "$env:ProgramFiles(x86)\nodejs\node.exe"
)

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Configuracao do servidor Windows" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$NodeCommand = Get-Command node -ErrorAction SilentlyContinue
if ($NodeCommand) {
    $NodePath = $NodeCommand.Source
} else {
    $NodePath = $NodeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

if (-not $NodePath) {
    Write-Host "Node.js nao encontrado." -ForegroundColor Yellow
    Write-Host "Voce pode instalar o Node.js LTS ou extrair um Node portatil em uma destas pastas:" -ForegroundColor Yellow
    Write-Host "  $Projeto\\node-portable" -ForegroundColor White
    Write-Host "  $Projeto\\nodejs" -ForegroundColor White
    Write-Host "  $Projeto\\tools\\node" -ForegroundColor White
    exit 1
}

if (-not (Test-Path $Launcher)) {
    throw "Script de inicializacao nao encontrado: $Launcher"
}

$TaskName = "SistemaComprasNode"

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

$PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Action = New-ScheduledTaskAction `
    -Execute $PowerShellExe `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Launcher`"" `
    -WorkingDirectory $Projeto

$Trigger = New-ScheduledTaskTrigger -AtStartup
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
    -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Inicia o servidor Node/Express do Sistema de Compras no boot do Windows." | Out-Null

Write-Host "Tarefa '$TaskName' criada com sucesso." -ForegroundColor Green
Write-Host "O servidor sera iniciado automaticamente a cada boot." -ForegroundColor Green
Write-Host ""
Write-Host "Para iniciar agora sem reiniciar, execute:" -ForegroundColor White
Write-Host "powershell -ExecutionPolicy Bypass -File `"$Launcher`"" -ForegroundColor White
