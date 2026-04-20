function Get-AudioSessions {
    $sessions = @()
    try {
        $enumerator = [AudioManager.AudioManagerHelper]::CreateDeviceEnumerator()
        $defaultDev = $null
        $enumerator.GetDefaultAudioEndpoint(
            [AudioManager.EDataFlow]::eRender,
            [AudioManager.ERole]::eConsole,
            [ref]$defaultDev
        ) | Out-Null

        $mgr2Iid = [AudioManager.AudioManagerHelper]::IID_IAudioSessionManager2
        $mgr2Obj = $null
        $defaultDev.Activate([ref]$mgr2Iid, 23, [IntPtr]::Zero, [ref]$mgr2Obj) | Out-Null
        $mgr2 = [AudioManager.IAudioSessionManager2]$mgr2Obj

        $sessionEnum = $null
        $mgr2.GetSessionEnumerator([ref]$sessionEnum) | Out-Null

        $count = 0
        $sessionEnum.GetCount([ref]$count) | Out-Null

        for ($i = 0; $i -lt $count; $i++) {
            $ctrl = $null
            $sessionEnum.GetSession($i, [ref]$ctrl) | Out-Null

            try {
                $ctrl2    = [AudioManager.IAudioSessionControl2]$ctrl
                $simpleVol = [AudioManager.ISimpleAudioVolume]$ctrl

                $pid = 0
                $ctrl2.GetProcessId([ref]$pid) | Out-Null

                # Skip the system sounds session (PID 0)
                if ($pid -eq 0) { continue }

                $displayName = ""
                $ctrl2.GetDisplayName([ref]$displayName) | Out-Null

                $proc = $null
                if ([string]::IsNullOrEmpty($displayName)) {
                    $proc = Get-Process -Id $pid -ErrorAction SilentlyContinue
                    $displayName = if ($proc -and $proc.MainWindowTitle) {
                        $proc.MainWindowTitle
                    } elseif ($proc) {
                        $proc.ProcessName
                    } else {
                        "PID $pid"
                    }
                }

                $vol   = 0.0
                $muted = $false
                $simpleVol.GetMasterVolume([ref]$vol)   | Out-Null
                $simpleVol.GetMute([ref]$muted)          | Out-Null

                $state = [AudioManager.AudioSessionState]::Inactive
                $ctrl.GetState([ref]$state) | Out-Null

                $sessions += [PSCustomObject]@{
                    SessionKey    = "$pid-$i"
                    Name          = $displayName
                    ProcessId     = $pid
                    PidLabel      = "PID: $pid"
                    VolumePercent = [math]::Round($vol * 100)
                    VolumeLabel   = "$([math]::Round($vol * 100))%"
                    IsMuted       = $muted
                    State         = $state
                    Icon          = Get-ProcessIcon -ProcessId $pid
                    SimpleVolume  = $simpleVol
                    SessionControl = $ctrl
                }
            } catch {
                # Session may have expired, skip it
            }
        }
    } catch {
        Write-Warning "Get-AudioSessions error: $_"
    }
    return $sessions
}

function Get-ProcessIcon {
    param([uint32]$ProcessId)
    $map = @{
        'chrome'            = '[web]'
        'firefox'           = '[web]'
        'msedge'            = '[web]'
        'spotify'           = '[music]'
        'discord'           = '[chat]'
        'slack'             = '[chat]'
        'teams'             = '[meet]'
        'zoom'              = '[meet]'
        'vlc'               = '[video]'
        'obs64'             = '[stream]'
        'obs32'             = '[stream]'
        'steam'             = '[game]'
        'epicgameslauncher' = '[game]'
        'mpc-hc64'          = '[video]'
        'foobar2000'        = '[music]'
        'itunes'            = '[music]'
        'winamp'            = '[music]'
    }
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($proc) {
            $key = $proc.ProcessName.ToLower()
            if ($map.ContainsKey($key)) { return $map[$key] }
        }
    } catch {}
    return '[audio]'
}
