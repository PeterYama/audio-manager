function Remove-AudioProfile {
    param([Parameter(Mandatory)]$Profile)
    try {
        $existing = Get-AudioProfiles
        $existing  = @($existing | Where-Object { $_.name -ne $Profile.name })
        $existing | ConvertTo-Json -Depth 5 | Set-Content -Path $sync.ProfilesPath -Encoding UTF8
        $sync.Profiles = Get-AudioProfiles
        Set-WPFStatus "Profile '$($Profile.name)' deleted."
    } catch {
        Write-Warning "Remove-AudioProfile error: $_"
    }
}
