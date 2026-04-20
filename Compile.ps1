#Requires -Version 5.1
<#
.SYNOPSIS
    Compiles all src/ files into a single AudioManager.ps1

.EXAMPLE
    .\Compile.ps1
    .\Compile.ps1 -OutFile "dist\AudioManager.ps1"
#>
param(
    [string]$OutFile = "AudioManager.ps1"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root    = $PSScriptRoot
$version = Get-Date -Format 'yy.MM.dd'

Write-Host "Audio Manager Compiler" -ForegroundColor Cyan
Write-Host "Version: $version" -ForegroundColor Gray
Write-Host ""

$output = [System.Text.StringBuilder]::new()

function Append {
    param([string]$text)
    [void]$output.AppendLine($text)
}

function AppendFile {
    param([string]$path, [string]$label = "")
    if (-not (Test-Path $path)) {
        Write-Warning "Missing file: $path"
        return
    }
    if ($label) { Write-Host "  + $label" -ForegroundColor DarkGray }
    $content = Get-Content -Path $path -Raw -Encoding UTF8
    [void]$output.AppendLine($content)
    [void]$output.AppendLine("")
}

# ─── 1. Header banner ─────────────────────────────────────────────────────────

Append @"
#Requires -Version 5.1
# ╔══════════════════════════════════════════════════════════════════╗
# ║           Audio Manager v$version — by PeterYama                ║
# ║   iwr -useb https://raw.githubusercontent.com/PeterYama/        ║
# ║       audio-manager/main/AudioManager.ps1 | iex                 ║
# ║   https://github.com/PeterYama/audio-manager                    ║
# ╚══════════════════════════════════════════════════════════════════╝
"@

# ─── 2. start.ps1 (with version injected) ─────────────────────────────────────

Write-Host "Building: scripts/start.ps1" -ForegroundColor Yellow
$startContent = (Get-Content "$root\src\scripts\start.ps1" -Raw -Encoding UTF8) `
    -replace '#\{VERSION\}#', $version
[void]$output.AppendLine($startContent)

# ─── 3. Embed CoreAudio.cs as a PowerShell variable ───────────────────────────

Write-Host "Building: core/CoreAudio.cs" -ForegroundColor Yellow
$csContent = Get-Content "$root\src\core\CoreAudio.cs" -Raw -Encoding UTF8
[void]$output.AppendLine('$script:CoreAudioCSharp = @' + "'" + "`n$csContent`n'@")
[void]$output.AppendLine("")

# ─── 4. Initialize-CoreAudio.ps1 ──────────────────────────────────────────────

Write-Host "Building: core/Initialize-CoreAudio.ps1" -ForegroundColor Yellow
AppendFile "$root\src\core\Initialize-CoreAudio.ps1"

# ─── 5. Private functions ──────────────────────────────────────────────────────

Write-Host "Building: functions/private/" -ForegroundColor Yellow
Get-ChildItem "$root\src\functions\private" -Filter "*.ps1" | Sort-Object Name | ForEach-Object {
    AppendFile $_.FullName $_.Name
}

# ─── 6. Public functions ───────────────────────────────────────────────────────

Write-Host "Building: functions/public/" -ForegroundColor Yellow
Get-ChildItem "$root\src\functions\public" -Filter "*.ps1" | Sort-Object Name | ForEach-Object {
    AppendFile $_.FullName $_.Name
}

# ─── 7. Embed XAML ────────────────────────────────────────────────────────────

Write-Host "Building: ui/MainWindow.xaml" -ForegroundColor Yellow
$xamlContent = Get-Content "$root\src\ui\MainWindow.xaml" -Raw -Encoding UTF8
[void]$output.AppendLine('$inputXML = @' + "'" + "`n$xamlContent`n'@")
[void]$output.AppendLine("")

# ─── 8. main.ps1 ──────────────────────────────────────────────────────────────

Write-Host "Building: scripts/main.ps1" -ForegroundColor Yellow
AppendFile "$root\src\scripts\main.ps1"

# ─── Write output ─────────────────────────────────────────────────────────────

$outPath = Join-Path $root $OutFile
$output.ToString() | Set-Content -Path $outPath -Encoding UTF8

Write-Host ""
Write-Host "Compiled -> $outPath ($([math]::Round((Get-Item $outPath).Length / 1KB, 1)) KB)" -ForegroundColor Green

# ─── Syntax validation ────────────────────────────────────────────────────────

Write-Host "Validating syntax..." -ForegroundColor Gray
$errors = $null
$tokens = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    $outPath, [ref]$tokens, [ref]$errors
)

if ($errors.Count -gt 0) {
    Write-Host "Syntax errors found:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  Line $($_.Extent.StartLineNumber): $($_.Message)" -ForegroundColor Red }
    exit 1
}

Write-Host "Syntax OK" -ForegroundColor Green
Write-Host ""
Write-Host "Run with:" -ForegroundColor Cyan
Write-Host "  .\AudioManager.ps1" -ForegroundColor White
