function Initialize-ApplicationsTab {
    $sync.WPFAppSessionList.Children.Clear()

    $sessions = $sync.AudioSessions
    if (-not $sessions -or $sessions.Count -eq 0) {
        $sync.WPFAppCount.Text = "No active audio sessions found."
        return
    }

    $sync.WPFAppCount.Text = "$($sessions.Count) session(s)"

    $bgColor  = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString("#16213E")
    $fgMain   = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString("#E0E0E0")
    $fgMuted  = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString("#8892B0")

    foreach ($session in ($sessions | Sort-Object Name)) {
        # --- outer card border ---
        $border = [System.Windows.Controls.Border]::new()
        $border.Background   = $bgColor
        $border.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $border.Padding      = [System.Windows.Thickness]::new(12, 8, 12, 8)
        $border.Margin       = [System.Windows.Thickness]::new(0, 0, 0, 8)

        # --- 5-column grid ---
        $grid = [System.Windows.Controls.Grid]::new()
        foreach ($w in @(36, 180, -1, 50, 70)) {
            $cd = [System.Windows.Controls.ColumnDefinition]::new()
            $cd.Width = if ($w -lt 0) { [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) } `
                        else           { [System.Windows.GridLength]::new($w) }
            $grid.ColumnDefinitions.Add($cd)
        }

        # Col 0: icon
        $iconTb = [System.Windows.Controls.TextBlock]::new()
        $iconTb.Text                = $session.Icon
        $iconTb.FontSize            = 14
        $iconTb.VerticalAlignment   = "Center"
        $iconTb.HorizontalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($iconTb, 0)

        # Col 1: name + pid
        $namePanel = [System.Windows.Controls.StackPanel]::new()
        $namePanel.VerticalAlignment = "Center"
        $namePanel.Margin = [System.Windows.Thickness]::new(8, 0, 0, 0)

        $nameTb = [System.Windows.Controls.TextBlock]::new()
        $nameTb.Text            = $session.Name
        $nameTb.FontWeight      = "SemiBold"
        $nameTb.Foreground      = $fgMain
        $nameTb.TextTrimming    = "CharacterEllipsis"

        $pidTb = [System.Windows.Controls.TextBlock]::new()
        $pidTb.Text       = $session.PidLabel
        $pidTb.Foreground = $fgMuted
        $pidTb.FontSize   = 10

        $namePanel.Children.Add($nameTb) | Out-Null
        $namePanel.Children.Add($pidTb)  | Out-Null
        [System.Windows.Controls.Grid]::SetColumn($namePanel, 1)

        # Col 2: volume slider
        $slider = [System.Windows.Controls.Slider]::new()
        $slider.Minimum           = 0
        $slider.Maximum           = 100
        $slider.Value             = $session.VolumePercent
        $slider.VerticalAlignment = "Center"
        $slider.Margin            = [System.Windows.Thickness]::new(8, 0, 8, 0)
        $slider.Tag               = [int]$session.ProcessId
        [System.Windows.Controls.Grid]::SetColumn($slider, 2)

        # Col 3: volume label (updated live by slider)
        $volLabel = [System.Windows.Controls.TextBlock]::new()
        $volLabel.Text                = $session.VolumeLabel
        $volLabel.Foreground          = $fgMain
        $volLabel.FontWeight          = "SemiBold"
        $volLabel.VerticalAlignment   = "Center"
        $volLabel.HorizontalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($volLabel, 3)

        # Col 4: mute toggle
        $muteBtn = [System.Windows.Controls.Primitives.ToggleButton]::new()
        $muteBtn.Content          = if ($session.IsMuted) { "[X]" } else { "Mute" }
        $muteBtn.IsChecked        = $session.IsMuted
        $muteBtn.VerticalAlignment = "Center"
        $muteBtn.Tag              = [int]$session.ProcessId
        $muteBtn.Padding          = [System.Windows.Thickness]::new(8, 4, 8, 4)
        [System.Windows.Controls.Grid]::SetColumn($muteBtn, 4)

        # --- events ---
        $capturedLabel = $volLabel
        $slider.Add_ValueChanged({
            $capturedLabel.Text = "$([math]::Round($this.Value))%"
        }.GetNewClosure())

        $slider.Add_PreviewMouseLeftButtonUp({
            $pid   = [int]$this.Tag
            $level = [float]($this.Value / 100.0)
            $sync._tmpPid   = $pid
            $sync._tmpLevel = $level
            Invoke-AudioManagerRunspace { Set-AppVolume -ProcessId $sync._tmpPid -Level $sync._tmpLevel }
        }.GetNewClosure())

        $muteBtn.Add_Click({
            $pid   = [int]$this.Tag
            $muted = [bool]$this.IsChecked
            $this.Content = if ($muted) { "[X]" } else { "Mute" }
            $sync._tmpPid   = $pid
            $sync._tmpMuted = $muted
            Invoke-AudioManagerRunspace { Set-AppMute -ProcessId $sync._tmpPid -Muted $sync._tmpMuted }
        }.GetNewClosure())

        # --- assemble ---
        $grid.Children.Add($iconTb)    | Out-Null
        $grid.Children.Add($namePanel) | Out-Null
        $grid.Children.Add($slider)    | Out-Null
        $grid.Children.Add($volLabel)  | Out-Null
        $grid.Children.Add($muteBtn)   | Out-Null
        $border.Child = $grid
        $sync.WPFAppSessionList.Children.Add($border) | Out-Null
    }
}

function Update-ApplicationsTab {
    Initialize-ApplicationsTab
}
