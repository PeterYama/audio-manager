function Get-AudioEnhancement {
    param([Parameter(Mandatory)][string]$DeviceId)
    # DeviceId format: {0.0.0.00000000}.{GUID}  - extract the GUID portion
    $guidPart = ($DeviceId -split '\.')[-1].Trim('{', '}')
    $regPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{$guidPart}\FxProperties"

    if (-not (Test-Path $regPath)) {
        # Try capture path
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{$guidPart}\FxProperties"
    }

    try {
        $val = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        # PKEY_AudioEndpoint_Disable_SysFx = {1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5
        $propName = "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5"
        if ($null -ne $val -and $null -ne $val.$propName) {
            # Value 0 = enhancements enabled, 1 = disabled
            return ($val.$propName -eq 0)
        }
        return $true  # Default: enhancements enabled
    } catch {
        return $true
    }
}
