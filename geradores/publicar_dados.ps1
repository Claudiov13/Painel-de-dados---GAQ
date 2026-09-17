param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'publicacao_lib.ps1')
$Root = (Resolve-Path -LiteralPath $Root).Path
$mutexName = 'Local\PainelGAQ_' + (Get-PainelHash $Root).Substring(0, 20)
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
$locked = $false
try {
    try { $locked = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $locked = $true }
    if (-not $locked) { throw 'Outra publicacao esta em andamento nesta maquina. Tente novamente.' }
    $defs = @(
        @{key='dados'; file='dados.js'; global='__PAINEL_DADOS__'},
        @{key='tags'; file='tags.js'; global='__PAINEL_TAGS__'},
        @{key='servicedesk'; file='chamados_servicedesk.js'; global='__SERVICE_DESK__'},
        @{key='fracionamento'; file='base_fracionamento.js'; global='__FRACIONAMENTO__'}
    )
    $sources = [ordered]@{}; $outputs = @()
    # Validate all inputs before changing any published source.
    foreach ($def in $defs) {
        $path = Join-Path $Root $def.file
        $raw = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
        $parts = $raw -split '\r?\nwindow\.__PAINEL_SOURCE_META__\s*=\s*', 2
        $json = ($parts[0] -replace ('^\s*window\.' + [regex]::Escape($def.global) + '\s*=\s*'), '').Trim() -replace ';\s*$', ''
        $value = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        $generated = $null
        if ($parts.Count -gt 1) {
            $oldMeta = ConvertFrom-Json -InputObject ($parts[1].Trim() -replace ';\s*$', '') -ErrorAction Stop
            $generated = $oldMeta.($def.key).generatedAt
        }
        if ($def.key -eq 'dados') {
            if ($value -isnot [Array] -or $value.Count -eq 0) { throw 'Base principal vazia ou invalida.' }
            foreach ($row in $value) { if ($row -isnot [pscustomobject]) { throw 'Registro principal invalido.' } }
            $count = $value.Count
        } elseif ($def.key -eq 'fracionamento') {
            if ($value.itens -isnot [Array]) { throw 'Base MXM invalida.' }; $count = $value.itens.Count
        } else {
            if ($value -isnot [pscustomobject]) { throw ('Objeto invalido: ' + $def.file) }
            $count = @($value.PSObject.Properties).Count
        }
        $revision = Get-PainelHash $json
        $meta = [ordered]@{revision=$revision; generatedAt=$generated; count=$count}
        $wrapper = @{}; $wrapper[$def.key] = $meta
        $content = 'window.' + $def.global + ' = ' + $json + ";`nwindow.__PAINEL_SOURCE_META__ = " + ($wrapper | ConvertTo-Json -Depth 5 -Compress) + ";`n"
        $outputs += @{path=$path; content=$content; original=$raw}
        $sources[$def.key] = [ordered]@{file=$def.file; revision=$revision; count=$count; generatedAt=$generated}
    }
    $revision = Get-PainelHash ($sources | ConvertTo-Json -Depth 6 -Compress)
    $manifestPath = Join-Path $Root 'publicacao.js'
    $manifest = [ordered]@{schemaVersion=1; revision=$revision; publishedAt=[DateTime]::UtcNow.ToString('o'); sources=$sources}
    foreach ($output in $outputs) { if ($output.content -cne $output.original) { Write-PainelAtomic $output.path $output.content } }
    # The manifest is committed last. Clients reject a partially synced set.
    Write-PainelAtomic $manifestPath ('window.__PAINEL_MANIFEST__ = ' + ($manifest | ConvertTo-Json -Depth 6 -Compress) + ";`n")
    Write-Host ('Publicacao validada: ' + $revision.Substring(0,12) + ' (' + $sources.dados.count + ' processos)')
} finally {
    if ($locked) { $mutex.ReleaseMutex() }; $mutex.Dispose()
}
