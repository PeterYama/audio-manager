function Get-AudioDevices {
    $result = @{ Render = @(); Capture = @() }
    try {
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        foreach ($flow in @([AudioManager.EDataFlow]::eRender, [AudioManager.EDataFlow]::eCapture)) {
            $collection = $null
            $enumerator.EnumAudioEndpoints($flow, [AudioManager.DeviceState]::Active, [ref]$collection) | Out-Null
            $count = 0
            $collection.GetCount([ref]$count) | Out-Null
            $devices = @()
            for ($i = 0; $i -lt $count; $i++) {
                $device = $null
                $collection.Item($i, [ref]$device) | Out-Null

                $id = ""
                $device.GetId([ref]$id) | Out-Null

                $store = $null
                $device.OpenPropertyStore(0, [ref]$store) | Out-Null  # STGM_READ = 0
                $key = [AudioManager.AudioManagerHelper]::PKEY_Device_FriendlyName()
                $pv  = New-Object AudioManager.PropVariant
                $store.GetValue([ref]$key, [ref]$pv) | Out-Null
                $name = $pv.GetStringValue()
                if ([string]::IsNullOrEmpty($name)) { $name = "Unknown Device" }

                # Get endpoint volume interface
                $volIid = [AudioManager.AudioManagerHelper]::IID_IAudioEndpointVolume
                $volObj = $null
                $device.Activate([ref]$volIid, 23, [IntPtr]::Zero, [ref]$volObj) | Out-Null
                $vol = [AudioManager.IAudioEndpointVolume]$volObj

                $level = 0.0
                $vol.GetMasterVolumeLevelScalar([ref]$level) | Out-Null
                $muted = $false
                $vol.GetMute([ref]$muted) | Out-Null

                $devices += [PSCustomObject]@{
                    DeviceId      = $id
                    Name          = $name
                    Flow          = $flow
                    VolumeScalar  = $level
                    IsMuted       = $muted
                    VolumeInterface = $vol
                }
            }
            if ($flow -eq [AudioManager.EDataFlow]::eRender) {
                $result.Render = $devices
            } else {
                $result.Capture = $devices
            }
        }
    } catch {
        Write-Warning "Get-AudioDevices error: $_"
    }
    return $result
}
