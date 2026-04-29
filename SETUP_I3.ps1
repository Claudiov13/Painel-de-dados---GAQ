# ============================================================
#  SETUP_I3.ps1 — Configuração do servidor dedicado
#  Execute como Administrador na máquina i3
# ============================================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Setup do Servidor Dedicado - Sistema GAQ " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Pasta do projeto (OneDrive sincronizado)
$PROJETO = Split-Path -Parent $MyInvocation.MyCommand.Path

# ── PASSO 1: Node.js ────────────────────────────────────────
Write-Host "[1/4] Verificando Node.js..." -ForegroundColor Yellow

$nodeOk = $false
try {
    $nodeVer = & node --version 2>&1
    if ($nodeVer -match "v\d+") {
        Write-Host "      Node.js já instalado: $nodeVer" -ForegroundColor Green
        $nodeOk = $true
    }
} catch {}

if (-not $nodeOk) {
    Write-Host "      Instalando Node.js LTS via winget..." -ForegroundColor Yellow
    winget install OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
    # Atualizar PATH na sessão atual
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
    Write-Host "      Node.js instalado." -ForegroundColor Green
}

# ── PASSO 2: Dependências do projeto ────────────────────────
Write-Host ""
Write-Host "[2/4] Instalando dependências do projeto (Express)..." -ForegroundColor Yellow
Set-Location $PROJETO
npm install
Write-Host "      Dependências OK." -ForegroundColor Green

# ── PASSO 3: PostgreSQL ─────────────────────────────────────
Write-Host ""
Write-Host "[3/4] Verificando PostgreSQL..." -ForegroundColor Yellow

$pgOk = $false
try {
    $pgVer = & psql --version 2>&1
    if ($pgVer -match "PostgreSQL") {
        Write-Host "      PostgreSQL já instalado: $pgVer" -ForegroundColor Green
        $pgOk = $true
    }
} catch {}

if (-not $pgOk) {
    Write-Host "      Instalando PostgreSQL 16 via winget..." -ForegroundColor Yellow
    winget install PostgreSQL.PostgreSQL.16 --accept-source-agreements --accept-package-agreements
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
    Write-Host "      PostgreSQL instalado." -ForegroundColor Green
}

# ── PASSO 4: Agendamento automático no boot ──────────────────
Write-Host ""
Write-Host "[4/4] Configurando inicialização automática no boot..." -ForegroundColor Yellow

# Encontra o node.exe instalado
$nodePath = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $nodePath) {
    # Tenta caminhos comuns
    $candidatos = @(
        "$env:ProgramFiles\nodejs\node.exe",
        "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
    )
    foreach ($c in $candidatos) {
        if (Test-Path $c) { $nodePath = $c; break }
    }
}

if ($nodePath) {
    $taskName = "SistemaComprasGAQ"
    $serverJs = Join-Path $PROJETO "server.js"

    # Remove tarefa anterior se existir
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

    $action  = New-ScheduledTaskAction -Execute $nodePath -Argument "`"$serverJs`"" -WorkingDirectory $PROJETO
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 0) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Servidor do Sistema de Compras GAQ - porta 3000"

    Write-Host "      Tarefa agendada: '$taskName' criada." -ForegroundColor Green
    Write-Host "      O servidor vai iniciar automaticamente a cada boot." -ForegroundColor Green
} else {
    Write-Host "      AVISO: node.exe não encontrado no PATH. Reinicie o PowerShell e rode este script novamente." -ForegroundColor Red
}

# ── Resultado final ──────────────────────────────────────────
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Setup concluído!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Descobre o IP da máquina
$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -eq "Dhcp" -or $_.PrefixOrigin -eq "Manual" } | Select-Object -First 1).IPAddress
Write-Host "  IP desta máquina: $ip" -ForegroundColor White
Write-Host "  Acesso ao sistema: http://${ip}:3000" -ForegroundColor White
Write-Host ""
Write-Host "  Próximo passo: pedir ao TI IP fixo para este endereço." -ForegroundColor Yellow
Write-Host ""

# Inicia o servidor imediatamente para testar
Write-Host "Iniciando servidor para teste..." -ForegroundColor Yellow
Start-Process -FilePath $nodePath -ArgumentList "`"$(Join-Path $PROJETO 'server.js')`"" -WorkingDirectory $PROJETO
Start-Sleep -Seconds 2
Write-Host "Servidor rodando em http://${ip}:3000" -ForegroundColor Green
Write-Host "Abra esse endereço no navegador desta máquina para confirmar." -ForegroundColor Green
Write-Host ""
Read-Host "Pressione Enter para fechar"
