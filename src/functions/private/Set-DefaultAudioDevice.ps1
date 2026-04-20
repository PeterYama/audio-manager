function Set-DefaultAudioDevice {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][AudioManager.EDataFlow]$Flow
    )
    try {
        $policy = [AudioManager.AudioManagerHelper]::CreatePolicyConfig()
        # Set for all three roles
        foreach ($role in @(
            [AudioManager.ERole]::eConsole,
            [AudioManager.ERole]::eMultimedia,
            [AudioManager.ERole]::eCommunications
        )) {
            $policy.SetDefaultEndpoint($DeviceId, $role) | Out-Null
        }
        return $true
    } catch {
        Write-Warning "Set-DefaultAudioDevice error: $_"
        return $false
    }
}
