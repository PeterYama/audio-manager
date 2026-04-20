function Invoke-WPFTab {
    param([Parameter(Mandatory)][string]$ClickedTab)

    $tabMap = @{
        WPFTab1BT = 'WPFTab1'
        WPFTab2BT = 'WPFTab2'
        WPFTab3BT = 'WPFTab3'
        WPFTab4BT = 'WPFTab4'
    }

    foreach ($btn in $tabMap.Keys) {
        $sync[$btn].IsChecked = ($btn -eq $ClickedTab)
    }

    $targetTab = $tabMap[$ClickedTab]
    foreach ($tab in $tabMap.Values) {
        $tabItem = $sync.WPFTabControl.Items | Where-Object { $_.Name -eq $tab }
        if ($tabItem) { $tabItem.IsSelected = ($tab -eq $targetTab) }
    }

    $sync.CurrentTab = $targetTab
}
