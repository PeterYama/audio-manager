$script:AMVersion = "#{VERSION}#"

# --- Elevation check ---

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin   = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    if ($PSCommandPath) {
        $target = $PSCommandPath
    } else {
        # Script was IEX'd - save to temp and relaunch elevated
        $target = "$env:TEMP\AudioManager_elevated.ps1"
        try {
            $scriptContent = (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/PeterYama/audio-manager/master/AudioManager.ps1" -UseBasicParsing)
            Set-Content -Path $target -Value $scriptContent -Encoding UTF8
        } catch {
            Write-Host "Could not auto-download for elevation. Please run PowerShell as Administrator." -ForegroundColor Red
            Start-Sleep -Seconds 4
            exit 1
        }
    }
    $relaunchArgs = "-ExecutionPolicy Bypass -File `"$target`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $relaunchArgs
    exit
}

# --- Assembly loading ---

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- Logging ---

$logDir = "$env:LOCALAPPDATA\AudioManager\logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
Start-Transcript -Path "$logDir\AudioManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" -Append -ErrorAction SilentlyContinue

# --- Profiles directory ---

$profilesDir = "$env:APPDATA\AudioManager"
if (-not (Test-Path $profilesDir)) { New-Item -ItemType Directory -Path $profilesDir -Force | Out-Null }

# --- Shared synchronized hashtable ---

$sync = [hashtable]::Synchronized(@{
    Form              = $null
    CurrentTab        = "Devices"
    RenderDevices     = @()
    CaptureDevices    = @()
    AudioSessions     = @()
    Profiles          = @()
    SelectedOutputId  = $null
    SelectedInputId   = $null
    IsRefreshing      = $false
    ProfilesPath      = "$env:APPDATA\AudioManager\profiles.json"
    Version           = $script:AMVersion
    NullGuid          = [guid]::Empty
})

# --- Runspace pool ---

$sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$sessionState.Variables.Add(
    [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('sync', $sync, '')
)
$sync.RunspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount, $sessionState, $Host)
$sync.RunspacePool.Open()
