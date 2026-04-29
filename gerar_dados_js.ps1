# =============================================================================
# gerar_dados_js.ps1
# Le "Base Consolidada GAQ IA.xlsx" e gera dados.js para o Painel GAQ.
#
# Como usar:
#   1. Execute manualmente via executar_agora.bat (duplo clique)
#   2. Ou agende via configurar_agendamento.ps1 (roda 07h e 19h todo dia)
# =============================================================================

# ── CAMINHOS ─────────────────────────────────────────────────────────────────

# Excel: pasta pessoal do OneDrive (KB_GAQ)
$EXCEL_PATH = Join-Path $env:USERPROFILE "OneDrive\KB_GAQ\Base Consolidada GAQ IA.xlsx"

# dados.js: mesma pasta deste script (pasta compartilhada "Sistema de compras")
# Usa $PSScriptRoot para evitar problemas com acentos no caminho
$OUTPUT_PATH = Join-Path $PSScriptRoot "dados.js"
$LOG_PATH    = Join-Path $PSScriptRoot "gerar_dados_js.log"

# ── LOG ───────────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Msg"
    try { Add-Content -Path $LOG_PATH -Value $line -Encoding UTF8 } catch {}
    Write-Host $line
}

# Mantém apenas as últimas 500 linhas do log
function Trim-Log {
    if (Test-Path $LOG_PATH) {
        $lines = Get-Content $LOG_PATH -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($lines -and $lines.Count -gt 500) {
            $lines | Select-Object -Last 500 | Set-Content $LOG_PATH -Encoding UTF8
        }
    }
}

# ── INICIO ────────────────────────────────────────────────────────────────────

Trim-Log
Write-Log "===== INICIO ====="
Write-Log "Excel  : $EXCEL_PATH"
Write-Log "Saida  : $OUTPUT_PATH"

# ── VERIFICACOES ──────────────────────────────────────────────────────────────

if (-not (Test-Path $EXCEL_PATH)) {
    Write-Log "Arquivo Excel nao encontrado: $EXCEL_PATH" "ERRO"
    Write-Log "Verifique se o OneDrive esta sincronizado." "ERRO"
    exit 1
}

# ── LEITURA DO EXCEL VIA COM ──────────────────────────────────────────────────

Write-Log "Abrindo Excel (somente leitura)..."

$xl = $null
$wb = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false
    $xl.ScreenUpdating = $false

    # Abre somente leitura (UpdateLinks=0, ReadOnly=$true)
    $wb = $xl.Workbooks.Open($EXCEL_PATH, 0, $true)
    $ws = $wb.Sheets.Item(1)

    Write-Log "Planilha aberta: '$($ws.Name)'"

    # Value2 retorna datas como serial Excel (numero inteiro) em vez de DateTime.
    # Ex.: 45587 = 22/10/2024. Exatamente o que o sistema JavaScript espera.
    $data = $ws.UsedRange.Value2

    if ($null -eq $data) {
        Write-Log "Planilha vazia." "ERRO"
        exit 1
    }

    # Para ranges de multiplas celulas: array 2D com indice base 1
    $rowCount = $data.GetUpperBound(0)
    $colCount = $data.GetUpperBound(1)
    Write-Log "Dimensoes: $rowCount linhas x $colCount colunas"

    # ── Cabecalhos (linha 1) ──────────────────────────────────────────────────
    $headers = @()
    for ($c = 1; $c -le $colCount; $c++) {
        $hVal = $data[1, $c]
        $headers += if ($null -ne $hVal) { $hVal.ToString().Trim() } else { "_col$c" }
    }

    # ── Dados (linhas 2 em diante) ────────────────────────────────────────────
    $records = [System.Collections.Generic.List[object]]::new()

    for ($r = 2; $r -le $rowCount; $r++) {
        # Ignora linhas totalmente vazias
        $hasValue = $false
        for ($c = 1; $c -le $colCount; $c++) {
            $v = $data[$r, $c]
            if ($null -ne $v -and $v.ToString().Trim() -ne "") { $hasValue = $true; break }
        }
        if (-not $hasValue) { continue }

        $rec = [ordered]@{}
        for ($c = 1; $c -le $colCount; $c++) {
            $key = $headers[$c - 1]
            $val = $data[$r, $c]

            if ($null -eq $val) {
                $rec[$key] = ""
            } elseif ($val -is [bool]) {
                # Boolean: preserva true/false (ex.: CPL_ENCONTRADO)
                $rec[$key] = $val
            } elseif ($val -is [double] -or $val -is [int] -or $val -is [long]) {
                # Numeros e seriais de data: converte double inteiro para int
                $num = [double]$val
                if ($num -eq [math]::Floor($num) -and $num -ge -2147483648 -and $num -le 2147483647) {
                    $rec[$key] = [int]$num
                } else {
                    $rec[$key] = $num
                }
            } else {
                $rec[$key] = $val.ToString().Trim()
            }
        }
        $records.Add($rec)
    }

    Write-Log "Registros lidos: $($records.Count)"

    if ($records.Count -eq 0) {
        Write-Log "Nenhum registro encontrado na planilha." "AVISO"
        exit 1
    }

    # ── GERACAO DO JSON ───────────────────────────────────────────────────────
    Write-Log "Gerando JSON..."

    $json = $records | ConvertTo-Json -Depth 5

    # PS 5.1 escapa caracteres nao-ASCII como \uXXXX no ConvertTo-Json.
    # Decodifica de volta para os caracteres reais (a, e, c, etc.)
    $json = [System.Text.RegularExpressions.Regex]::Replace(
        $json,
        '\\u([0-9a-fA-F]{4})',
        { param($m) [char][convert]::ToInt32($m.Groups[1].Value, 16) }
    )

    $content = "window.__PAINEL_DADOS__ = $json;`n"

    # Escreve UTF-8 sem BOM (compativel com <script src="./dados.js">)
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($OUTPUT_PATH, $content, $utf8NoBom)

    Write-Log "dados.js gerado: $($records.Count) registros"
    Write-Log "Caminho: $OUTPUT_PATH"

} catch {
    Write-Log "ERRO inesperado: $($_.Exception.Message)" "ERRO"
    Write-Log "Linha: $($_.InvocationInfo.ScriptLineNumber)" "ERRO"
    exit 1

} finally {
    if ($null -ne $wb)  { try { $wb.Close($false) } catch {} }
    if ($null -ne $xl)  {
        try { $xl.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Log "===== CONCLUIDO ====="
