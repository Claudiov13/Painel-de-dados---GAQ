# ASCII source for Windows PowerShell 5.1. No permanent backup files.
function Write-PainelAtomic {
    param([string]$Path, [string]$Content)
    $full = [System.IO.Path]::GetFullPath($Path)
    $temp = $full + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [System.IO.File]::WriteAllText($temp, $Content, (New-Object System.Text.UTF8Encoding($false)))
        if ([System.IO.File]::Exists($full)) { [System.IO.File]::Replace($temp, $full, [System.Management.Automation.Language.NullString]::Value) }
        else { [System.IO.File]::Move($temp, $full) }
    } finally {
        if ([System.IO.File]::Exists($temp)) { [System.IO.File]::Delete($temp) }
    }
}
function Get-PainelHash {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}
function Write-PainelSource {
    param([string]$Path, [string]$GlobalName, [string]$Key, [string]$Json)
    $null = ConvertFrom-Json -InputObject $Json -ErrorAction Stop
    $meta = @{}; $meta[$Key] = @{ generatedAt = [DateTime]::UtcNow.ToString('o') }
    $content = "window.$GlobalName = $Json;`nwindow.__PAINEL_SOURCE_META__ = " + ($meta | ConvertTo-Json -Compress) + ";`n"
    Write-PainelAtomic -Path $Path -Content $content
}
