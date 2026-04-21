# Parse XAML

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' `
                       -replace 'xmlns:d="[^"]*"', '' `
                       -replace 'xmlns:mc="[^"]*"', '' `
                       -replace "x:Class=`"[^`"]*`"", ''

# Use Parse() so that WPF registers element names into the namescope.
# XmlNodeReader skips namescope registration, making FindName() return
# null for every control even though the Window loads and renders fine.
$sync.Form = [Windows.Markup.XamlReader]::Parse($inputXML)

# Bind all named controls into $sync by walking the logical tree
# (Parse() registers names, so FindName works correctly here)

[xml]$xaml = $inputXML
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $ctrlName = $_.GetAttribute('Name')
    if ($ctrlName) {
        $ctrl = $sync.Form.FindName($ctrlName)
        if ($ctrl) { $sync[$ctrlName] = $ctrl }
    }
}

# Build runspace pool here (after all private/public functions are defined)
# so worker threads have every function available.
# Add-Type (CoreAudio) loads into the .NET AppDomain shared by all runspaces,
# so the C# types are automatically accessible without re-registering them.

$_iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$_iss.Variables.Add(
    [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('sync', $sync, '')
)

# Only inject our own functions - avoids enumerating hundreds of built-in PS functions
$_amPattern = '^(Get-Audio|Set-Audio|Get-Device|Set-Device|Get-App|Set-App|' +
              'Set-Default|Remove-Audio|Save-Audio|Restore-Audio|' +
              'Initialize-|Update-|Invoke-WPF|Set-WPFStatus|Invoke-Audio)'
Get-Command -CommandType Function |
    Where-Object { $_.Name -match $_amPattern } |
    ForEach-Object {
        try {
            $entry = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new(
                $_.Name, $_.ScriptBlock.ToString()
            )
            $_iss.Commands.Add($entry)
        } catch {}
    }

$sync.RunspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount, $_iss, $Host)
$sync.RunspacePool.Open()

# Version label

$sync.WPFVersionLabel.Text = "v$($sync.Version)"

# Tab navigation

foreach ($tabBtn in @('WPFTab1BT','WPFTab2BT','WPFTab3BT','WPFTab4BT')) {
    $btnName = $tabBtn
    $sync[$btnName].Add_Click({
        Invoke-WPFTab -ClickedTab $btnName
    }.GetNewClosure())
}

# Master volume slider

$sync.WPFMasterVolumeSlider.Add_ValueChanged({
    $pct = [math]::Round($sync.WPFMasterVolumeSlider.Value)
    $sync.WPFMasterVolumeLabel.Text = "$pct%"
})

$sync.WPFMasterVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFMasterVolumeSlider.Value / 100.0
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceVolume -DeviceId $sync.SelectedOutputId -Level $level }
    }
})

$sync.WPFMasterMuteButton.Add_Click({
    $muted = $sync.WPFMasterMuteButton.IsChecked
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceMute -DeviceId $sync.SelectedOutputId -Muted $muted }
    }
    $sync.WPFMasterMuteButton.Content = if ($muted) { "[X]" } else { "Mute" }
})

# Output device list selection

$sync.WPFOutputDeviceList.Add_SelectionChanged({
    $selected = $sync.WPFOutputDeviceList.SelectedItem
    if (-not $selected) {
        $sync.WPFSetDefaultOutput.IsEnabled   = $false
        $sync.WPFOutputVolumeSlider.IsEnabled = $false
        $sync.WPFOutputMuteButton.IsEnabled   = $false
        return
    }
    $sync.SelectedOutputId = $selected.Tag
    $sync.WPFSetDefaultOutput.IsEnabled   = $true
    $sync.WPFOutputVolumeSlider.IsEnabled = $true
    $sync.WPFOutputMuteButton.IsEnabled   = $true

    $dev = $sync.RenderDevices | Where-Object { $_.DeviceId -eq $selected.Tag } | Select-Object -First 1
    if ($dev) {
        $pct = [math]::Round($dev.VolumeScalar * 100)
        $sync.WPFOutputVolumeSlider.Value   = $pct
        $sync.WPFOutputVolumeLabel.Text     = "$pct%"
        $sync.WPFOutputMuteButton.IsChecked = $dev.IsMuted
        $sync.WPFMasterVolumeSlider.Value   = $pct
        $sync.WPFMasterVolumeLabel.Text     = "$pct%"
        $sync.WPFMasterMuteButton.IsChecked = $dev.IsMuted
        $sync.WPFMasterMuteButton.Content   = if ($dev.IsMuted) { "[X]" } else { "Mute" }
    }
})

$sync.WPFOutputVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFOutputVolumeSlider.Value / 100.0
    $pct   = [math]::Round($sync.WPFOutputVolumeSlider.Value)
    $sync.WPFOutputVolumeLabel.Text   = "$pct%"
    $sync.WPFMasterVolumeSlider.Value = $pct
    $sync.WPFMasterVolumeLabel.Text   = "$pct%"
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceVolume -DeviceId $sync.SelectedOutputId -Level $level }
    }
})

$sync.WPFOutputMuteButton.Add_Click({
    $muted = $sync.WPFOutputMuteButton.IsChecked
    if ($sync.SelectedOutputId) {
        Invoke-AudioManagerRunspace { Set-DeviceMute -DeviceId $sync.SelectedOutputId -Muted $muted }
    }
})

# Input device list selection

$sync.WPFInputDeviceList.Add_SelectionChanged({
    $selected = $sync.WPFInputDeviceList.SelectedItem
    if (-not $selected) {
        $sync.WPFSetDefaultInput.IsEnabled   = $false
        $sync.WPFInputVolumeSlider.IsEnabled = $false
        $sync.WPFInputMuteButton.IsEnabled   = $false
        return
    }
    $sync.SelectedInputId = $selected.Tag
    $sync.WPFSetDefaultInput.IsEnabled   = $true
    $sync.WPFInputVolumeSlider.IsEnabled = $true
    $sync.WPFInputMuteButton.IsEnabled   = $true

    $dev = $sync.CaptureDevices | Where-Object { $_.DeviceId -eq $selected.Tag } | Select-Object -First 1
    if ($dev) {
        $pct = [math]::Round($dev.VolumeScalar * 100)
        $sync.WPFInputVolumeSlider.Value   = $pct
        $sync.WPFInputVolumeLabel.Text     = "$pct%"
        $sync.WPFInputMuteButton.IsChecked = $dev.IsMuted
    }
})

$sync.WPFInputVolumeSlider.Add_PreviewMouseLeftButtonUp({
    $level = $sync.WPFInputVolumeSlider.Value / 100.0
    $pct   = [math]::Round($sync.WPFInputVolumeSlider.Value)
    $sync.WPFInputVolumeLabel.Text = "$pct%"
    if ($sync.SelectedInputId) {
        Invoke-AudioManagerRunspace { Set-DeviceVolume -DeviceId $sync.SelectedInputId -Level $level }
    }
})

$sync.WPFInputMuteButton.Add_Click({
    $muted = $sync.WPFInputMuteButton.IsChecked
    if ($sync.SelectedInputId) {
        Invoke-AudioManagerRunspace { Set-DeviceMute -DeviceId $sync.SelectedInputId -Muted $muted }
    }
})

# Formats tab device picker change

$sync.WPFFormatOutputDevice.Add_SelectionChanged({ Update-OutputFormatDisplay })
$sync.WPFFormatInputDevice.Add_SelectionChanged({  Update-InputFormatDisplay  })

# Enhancements toggles

$sync.WPFOutputEnhancementsToggle.Add_Click({
    $enabled  = $sync.WPFOutputEnhancementsToggle.IsChecked
    $selected = $sync.WPFFormatOutputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    Invoke-AudioManagerRunspace {
        $ok = Set-AudioEnhancement -DeviceId $deviceId -Enabled $enabled
        Invoke-WPFUIThread {
            $sync.WPFOutputEnhancementsToggle.Content = if ($enabled) { "Enhancements: ON" } else { "Enhancements: OFF" }
            Set-WPFStatus (if ($ok) { "Output enhancements updated." } else { "Failed to change enhancements." })
        }
    }
})

$sync.WPFInputEnhancementsToggle.Add_Click({
    $enabled  = $sync.WPFInputEnhancementsToggle.IsChecked
    $selected = $sync.WPFFormatInputDevice.SelectedItem
    if (-not $selected) { return }
    $deviceId = $selected.Tag
    Invoke-AudioManagerRunspace {
        $ok = Set-AudioEnhancement -DeviceId $deviceId -Enabled $enabled
        Invoke-WPFUIThread {
            $sync.WPFInputEnhancementsToggle.Content = if ($enabled) { "Enhancements: ON" } else { "Enhancements: OFF" }
            Set-WPFStatus (if ($ok) { "Input enhancements updated." } else { "Failed to change enhancements." })
        }
    }
})

# Profile list selection

$sync.WPFProfileList.Add_SelectionChanged({
    $selected   = $sync.WPFProfileList.SelectedItem
    $hasProfile = ($selected -and $selected.Tag)
    $sync.WPFRestoreProfile.IsEnabled = $hasProfile
    $sync.WPFDeleteProfile.IsEnabled  = $hasProfile
    if ($hasProfile) {
        $p    = $selected.Tag
        $info = "Created: $($p.created)"
        if ($p.defaultOutputDeviceName) { $info += "`nOutput: $($p.defaultOutputDeviceName)" }
        if ($p.defaultInputDeviceName)  { $info += "`nInput:  $($p.defaultInputDeviceName)" }
        $sync.WPFProfileInfo.Text = $info
    }
})

# Button dispatcher

foreach ($btnName in @(
    'WPFRefreshButton', 'WPFRefreshApps',
    'WPFSetDefaultOutput', 'WPFSetDefaultInput',
    'WPFApplyOutputFormat', 'WPFApplyInputFormat',
    'WPFSaveProfile', 'WPFRestoreProfile', 'WPFDeleteProfile'
)) {
    $name = $btnName
    if ($sync[$name]) {
        $sync[$name].Add_Click({
            Invoke-WPFButton -ClickedButton $name
        }.GetNewClosure())
    }
}

# Initial data load

Set-WPFStatus "Loading audio devices..."
Invoke-WPFRefreshAll

# Show window

$sync.Form.ShowDialog() | Out-Null

# Cleanup

$sync.RunspacePool.Close()
$sync.RunspacePool.Dispose()
Stop-Transcript -ErrorAction SilentlyContinue
