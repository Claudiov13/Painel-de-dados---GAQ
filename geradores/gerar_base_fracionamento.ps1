# =============================================================================
# gerar_base_fracionamento.ps1
# Le "Base de fracionamento MXM.xlsx" (aba unica) e gera base_fracionamento.js
# para o Painel GAQ.
#
# Saida: window.__FRACIONAMENTO__ = { gerado, fonte, totais, itens:[...] }
#
# Como usar:
#   powershell -ExecutionPolicy Bypass -File gerar_base_fracionamento.ps1
#
# Observacoes:
#   - Codigo-fonte mantido em ASCII puro de proposito: o Windows PowerShell 5.1
#     le .ps1 sem BOM como ANSI, entao acentos no fonte quebram o parser. Os
#     cabecalhos da planilha (que tem acentos) sao casados removendo acentos
#     em runtime, nao por literal no codigo.
#   - NAO depende do Excel instalado (le o XLSX direto como ZIP), entao funciona
#     mesmo com a planilha aberta.
# =============================================================================

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$EXCEL_FILE_NAME = "Base de fracionamento MXM.xlsx"

# ── Localiza o Excel (sem literais acentuados no fonte) ───────────────────────
$EXCEL_PATH = $null
$roots = @()
if ($env:OneDriveCommercial) { $roots += $env:OneDriveCommercial }
# Qualquer pasta "OneDrive*" no perfil cobre "OneDrive - Servico Social ..." com ou sem acento.
$roots += @(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })

foreach ($root in ($roots | Select-Object -Unique)) {
    if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }
    $kbDirs = @()
    $directKb = Join-Path $root "KB_GAQ"
    if (Test-Path -LiteralPath $directKb) { $kbDirs += $directKb }
    $personalDirs = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "claudio*Pessoal" }
    foreach ($dir in $personalDirs) {
        $kb = Join-Path $dir.FullName "KB_GAQ"
        if (Test-Path -LiteralPath $kb) { $kbDirs += $kb }
    }
    foreach ($kbDir in ($kbDirs | Select-Object -Unique)) {
        $candidate = Join-Path $kbDir $EXCEL_FILE_NAME
        if (Test-Path -LiteralPath $candidate) { $EXCEL_PATH = $candidate; break }
    }
    if ($EXCEL_PATH) { break }
}

# Saida na raiz do projeto (pasta acima de geradores/), onde o index.html le
$OUTPUT_PATH = Join-Path (Split-Path -Parent $PSScriptRoot) "base_fracionamento.js"

if (-not $EXCEL_PATH -or -not (Test-Path -LiteralPath $EXCEL_PATH)) {
    Write-Host "ERRO: '$EXCEL_FILE_NAME' nao encontrado em KB_GAQ. Verifique o OneDrive."
    exit 1
}
Write-Host "Excel : $EXCEL_PATH"
Write-Host "Saida : $OUTPUT_PATH"

# ── Helpers ──────────────────────────────────────────────────────────────────
function Strip-Accents([string]$s) {
    if (-not $s) { return '' }
    $n = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString()
}
function Get-ColLetter([string]$ref) { return ($ref -replace '[0-9]', '') }
function To-Num($s) {
    if ($null -eq $s -or "$s" -eq "") { return 0.0 }
    $n = 0.0
    if ([double]::TryParse([string]$s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return 0.0
}
function Norm-Enq($s) {
    $t = ("" + $s).Trim().ToUpper()
    if ($t -match 'DL\s*2|DL\s*II') { return 'DL 2' }
    if ($t -match 'DL\s*1|DL\s*I')  { return 'DL 1' }
    return $t
}

# ── Leitura do XLSX (ZIP, sem automacao do Excel) ────────────────────────────
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$fs  = [System.IO.File]::Open($EXCEL_PATH, 'Open', 'Read', 'ReadWrite')
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)

function Read-Entry($name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

try {
    $sst = @()
    $sstTxt = Read-Entry 'xl/sharedStrings.xml'
    if ($sstTxt) {
        [xml]$sx2 = $sstTxt
        foreach ($si in $sx2.sst.si) {
            if ($si.t) { $sst += [string]$si.t }
            elseif ($si.r) { $sst += (($si.r | ForEach-Object { $_.t }) -join '') }
            else { $sst += '' }
        }
    }

    [xml]$sx = Read-Entry 'xl/worksheets/sheet1.xml'
    $rows = @($sx.worksheet.sheetData.row)
    if ($rows.Count -lt 2) { throw "Planilha sem dados suficientes." }

    function Get-CellVal($c) {
        if ($c.t -eq 's') { return $sst[[int]$c.v] }
        elseif ($c.t -eq 'inlineStr') { return [string]$c.is.t }
        else { return [string]$c.v }
    }

    # A 1a linha do sheetData e o cabecalho (a planilha comeca em A2).
    # Monta lista (normHeader, colLetter) sem depender de acento no fonte.
    $hdrs = @()
    foreach ($c in $rows[0].c) {
        $raw = ("" + (Get-CellVal $c)).Trim()
        if ($raw) {
            $hdrs += [PSCustomObject]@{ Norm = (Strip-Accents $raw).ToLower().Trim(); Col = (Get-ColLetter $c.r) }
        }
    }
    function Find-Col([scriptblock]$pred) {
        $m = $hdrs | Where-Object { & $pred $_.Norm } | Select-Object -First 1
        if ($m) { return $m.Col } else { return $null }
    }

    $colCod    = Find-Col { param($h) $h -like 'cod*grupo*cota*' }
    $colDesc   = Find-Col { param($h) $h -like 'desc*grupo*cota*' }
    $colVlrT   = Find-Col { param($h) $h -eq 'soma de vlr total' }
    $colVlrA   = Find-Col { param($h) $h -like 'soma de vlr total atend*' -or $h -like '*atend*' }
    $colStatus = Find-Col { param($h) $h -eq 'status' }
    $colEnq    = Find-Col { param($h) $h -like 'enquadr*' }

    if (-not $colCod -or -not $colVlrT -or -not $colEnq) {
        throw ("Cabecalhos esperados nao encontrados. Achados: " + (($hdrs | ForEach-Object { $_.Norm }) -join ', '))
    }

    $itens = [System.Collections.Generic.List[object]]::new()
    for ($k = 1; $k -lt $rows.Count; $k++) {
        $h = @{}
        foreach ($c in $rows[$k].c) { $h[(Get-ColLetter $c.r)] = (Get-CellVal $c) }

        $cod  = ("" + $h[$colCod]).Trim()
        $desc = ("" + $h[$colDesc]).Trim()
        $enq  = Norm-Enq $h[$colEnq]
        if ($cod -eq "" -and $desc -eq "" -and $enq -eq "") { continue }

        $itens.Add([ordered]@{
            cod      = $cod
            desc     = $desc
            vlrTotal = [math]::Round((To-Num $h[$colVlrT]), 2)
            vlrAtend = [math]::Round((To-Num $h[$colVlrA]), 2)
            status   = ("" + $h[$colStatus]).Trim().ToUpper()
            dl       = $enq
        })
    }

    $dl1 = @($itens | Where-Object { $_.dl -eq 'DL 1' })
    $dl2 = @($itens | Where-Object { $_.dl -eq 'DL 2' })
    # itens sao OrderedDictionary; Measure-Object -Property nao le chaves, soma manual.
    $sumVlr = { param($arr) $s = 0.0; foreach ($it in $arr) { $s += [double]$it.vlrTotal }; [math]::Round($s, 2) }

    $payload = [ordered]@{
        gerado = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        fonte  = $EXCEL_FILE_NAME
        totais = [ordered]@{
            grupos = $itens.Count
            dl1 = [ordered]@{ grupos = $dl1.Count; vlrTotal = (& $sumVlr $dl1) }
            dl2 = [ordered]@{ grupos = $dl2.Count; vlrTotal = (& $sumVlr $dl2) }
        }
        itens = $itens
    }

    $json = $payload | ConvertTo-Json -Depth 6
    $json = [System.Text.RegularExpressions.Regex]::Replace(
        $json, '\\u([0-9a-fA-F]{4})',
        { param($m) [char][convert]::ToInt32($m.Groups[1].Value, 16) }
    )

    $content = "window.__FRACIONAMENTO__ = $json;`n"
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($OUTPUT_PATH, $content, $utf8NoBom)

    Write-Host ("OK: base_fracionamento.js gerado com {0} itens (DL1={1}, DL2={2})." -f $itens.Count, $dl1.Count, $dl2.Count)
}
finally {
    $zip.Dispose()
    $fs.Close()
}
