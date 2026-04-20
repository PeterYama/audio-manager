function Set-AudioEnhancement {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][bool]$Enabled
    )
    $guidPart = ($DeviceId -split '\.')[-1].Trim('{', '}')
    $regPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{$guidPart}\FxProperties"

    if (-not (Test-Path $regPath)) {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{$guidPart}\FxProperties"
    }

    try {
        $propName = "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5"
        $value    = if ($Enabled) { 0 } else { 1 }  # 0 = enabled, 1 = disabled

        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name $propName -Value $value -Type DWord -Force

        # Restart audio service to apply changes
        Restart-Service -Name audiosrv -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        Write-Warning "Set-AudioEnhancement error: $_"
        return $false
    }
}
