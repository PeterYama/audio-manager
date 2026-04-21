function Invoke-WPFButton {
    param([Parameter(Mandatory)][string]$ClickedButton)

    switch ($ClickedButton) {

        "WPFRefreshButton" {
            Invoke-WPFRefreshAll
        }

        "WPFRefreshApps" {
            Set-WPFStatus "Refreshing application sessions..."
            Invoke-AudioManagerRunspace {
                $sync.AudioSessions = Get-AudioSessions
                Invoke-WPFUIThread { Update-ApplicationsTab }
                Invoke-WPFUIThread { Set-WPFStatus "Found $($sync.AudioSessions.Count) active session(s)." }
            }
        }

        "WPFSetDefaultOutput" {
            $selected = $sync.WPFOutputDeviceList.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an output device first."; return }
            $deviceId = $selected.Tag
            $devName  = $selected.Content
            Set-WPFStatus "Setting default output to '$devName'..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DefaultAudioDevice -DeviceId $deviceId
                Invoke-WPFUIThread {
                    Set-WPFStatus (if ($ok) { "Default output set to '$devName'." } else { "Failed to set default output." })
                }
            }
        }

        "WPFSetDefaultInput" {
            $selected = $sync.WPFInputDeviceList.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an input device first."; return }
            $deviceId = $selected.Tag
            $devName  = $selected.Content
            Set-WPFStatus "Setting default input to '$devName'..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DefaultAudioDevice -DeviceId $deviceId
                Invoke-WPFUIThread {
                    Set-WPFStatus (if ($ok) { "Default input set to '$devName'." } else { "Failed to set default input." })
                }
            }
        }

        "WPFApplyOutputFormat" {
            $selected = $sync.WPFFormatOutputDevice.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an output device first."; return }
            $deviceId  = $selected.Tag
            $srItem    = $sync.WPFOutputSampleRate.SelectedItem
            $bdItem    = $sync.WPFOutputBitDepth.SelectedItem
            if (-not $srItem -or -not $bdItem) { Set-WPFStatus "Select sample rate and bit depth."; return }
            $sampleRate = [uint32]$srItem.Tag
            $bitDepth   = [uint16]$bdItem.Tag
            Set-WPFStatus "Applying output format: ${sampleRate} Hz / ${bitDepth}-bit..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DeviceFormat -DeviceId $deviceId -SampleRate $sampleRate -BitDepth $bitDepth
                Invoke-WPFUIThread {
                    if ($ok) {
                        Update-OutputFormatDisplay
                        Set-WPFStatus "Output format applied."
                    } else {
                        Set-WPFStatus "Failed to apply output format."
                    }
                }
            }
        }

        "WPFApplyInputFormat" {
            $selected = $sync.WPFFormatInputDevice.SelectedItem
            if (-not $selected) { Set-WPFStatus "Select an input device first."; return }
            $deviceId  = $selected.Tag
            $srItem    = $sync.WPFInputSampleRate.SelectedItem
            $bdItem    = $sync.WPFInputBitDepth.SelectedItem
            if (-not $srItem -or -not $bdItem) { Set-WPFStatus "Select sample rate and bit depth."; return }
            $sampleRate = [uint32]$srItem.Tag
            $bitDepth   = [uint16]$bdItem.Tag
            Set-WPFStatus "Applying input format: ${sampleRate} Hz / ${bitDepth}-bit..."
            Invoke-AudioManagerRunspace {
                $ok = Set-DeviceFormat -DeviceId $deviceId -SampleRate $sampleRate -BitDepth $bitDepth
                Invoke-WPFUIThread {
                    if ($ok) {
                        Update-InputFormatDisplay
                        Set-WPFStatus "Input format applied."
                    } else {
                        Set-WPFStatus "Failed to apply input format."
                    }
                }
            }
        }

        "WPFSaveProfile" {
            $name = $sync.WPFProfileNameInput.Text
            Save-AudioProfile -Name $name
            Invoke-WPFUIThread { Update-ProfilesTab }
        }

        "WPFRestoreProfile" {
            $selected = $sync.WPFProfileList.SelectedItem
            if (-not $selected -or -not $selected.Tag) { Set-WPFStatus "Select a profile to restore."; return }
            $profile = $selected.Tag
            Set-WPFStatus "Restoring profile '$($profile.name)'..."
            Invoke-AudioManagerRunspace {
                $ok = Restore-AudioProfile -Profile $profile
                if ($ok) {
                    # Re-read devices after potential device switch
                    $devices = Get-AudioDevices
                    $sync.RenderDevices  = $devices.Render
                    $sync.CaptureDevices = $devices.Capture
                    Invoke-WPFUIThread { Invoke-WPFRefreshAll }
                }
            }
        }

        "WPFDeleteProfile" {
            $selected = $sync.WPFProfileList.SelectedItem
            if (-not $selected -or -not $selected.Tag) { Set-WPFStatus "Select a profile to delete."; return }
            Remove-AudioProfile -Profile $selected.Tag
            Invoke-WPFUIThread { Update-ProfilesTab }
        }
    }
}
