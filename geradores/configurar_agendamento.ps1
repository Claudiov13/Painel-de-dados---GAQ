# =============================================================================
# configurar_agendamento.ps1
# Cria duas tarefas no Agendador de Tarefas do Windows:
#   - 07:00 todos os dias (atualizacao matinal)
#   - 19:00 todos os dias (atualizacao noturna)
#
# Execute este script UMA UNICA VEZ (duplo clique ou como Administrador).
# =============================================================================

# O script principal esta na mesma pasta que este arquivo
$PS_SCRIPT  = Join-Path $PSScriptRoot "gerar_dados_js.ps1"
$PS_EXE     = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$ARGS_STR   = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PS_SCRIPT`""
$TASK_NAME  = "Painel GAQ - Atualizar dados.js"

Write-Host ""
Write-Host " Configurando Agendador de Tarefas do Windows..." -ForegroundColor Cyan
Write-Host " Script: $PS_SCRIPT"
Write-Host ""

if (-not (Test-Path $PS_SCRIPT)) {
    Write-Host " [ERRO] Script nao encontrado: $PS_SCRIPT" -ForegroundColor Red
    Write-Host " Verifique se gerar_dados_js.ps1 esta na mesma pasta."
    pause
    exit 1
}

# Remove tarefa anterior com mesmo nome, se existir
$existing = Get-ScheduledTask -TaskName $TASK_NAME -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host " Removendo tarefa anterior..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TASK_NAME -Confirm:$false
}

# Acao: executar o script PowerShell
$action = New-ScheduledTaskAction -Execute $PS_EXE -Argument $ARGS_STR

# Gatilhos: 07:00 e 19:00 diariamente
$trigger1 = New-ScheduledTaskTrigger -Daily -At "07:00"
$trigger2 = New-ScheduledTaskTrigger -Daily -At "19:00"

# Configuracoes
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -RestartCount 1 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false

# Executa como usuario atual (nao precisa de senha)
$principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

try {
    Register-ScheduledTask `
        -TaskName    $TASK_NAME `
        -Action      $action `
        -Trigger     $trigger1, $trigger2 `
        -Settings    $settings `
        -Principal   $principal `
        -Description "Atualiza dados.js do Painel GAQ lendo a Base Consolidada GAQ IA.xlsx" `
        | Out-Null

    Write-Host " [OK] Tarefa '$TASK_NAME' criada!" -ForegroundColor Green
    Write-Host ""
    Write-Host "      Execucao: todos os dias as 07:00 e 19:00" -ForegroundColor White
    Write-Host ""
    Write-Host " Para verificar ou ajustar horarios:"
    Write-Host " Painel de Controle > Ferramentas Administrativas > Agendador de Tarefas"
    Write-Host " A tarefa aparece em: Biblioteca do Agendador de Tarefas"
    Write-Host ""

    # Oferecer execucao imediata para teste (roda o script diretamente, sem depender da tarefa)
    $runNow = Read-Host " Deseja executar agora para testar? (S/N)"
    if ($runNow -ieq "S") {
        Write-Host ""
        Write-Host " Executando..." -ForegroundColor Cyan
        & powershell -NoProfile -ExecutionPolicy Bypass -File $PS_SCRIPT
        if ($LASTEXITCODE -eq 0) {
            Write-Host " [OK] Executado com sucesso!" -ForegroundColor Green
        } else {
            Write-Host " [AVISO] Verifique o log: $PSScriptRoot\gerar_dados_js.log" -ForegroundColor Yellow
        }
    }

} catch {
    Write-Host " [ERRO] Falha ao criar a tarefa: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host " Tente executar este script como Administrador:" -ForegroundColor Yellow
    Write-Host " Clique direito no arquivo > Executar como administrador"
}

Write-Host ""
Write-Host " Pressione qualquer tecla para fechar..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
