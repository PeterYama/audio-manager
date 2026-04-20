function Get-AudioProfiles {
    try {
        if (Test-Path $sync.ProfilesPath) {
            $json = Get-Content -Path $sync.ProfilesPath -Raw -ErrorAction Stop
            $profiles = $json | ConvertFrom-Json
            return @($profiles)
        }
    } catch {
        Write-Warning "Get-AudioProfiles error: $_"
    }
    return @()
}
