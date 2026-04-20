function Invoke-WPFUIThread {
    param([Parameter(Mandatory)][scriptblock]$Code)
    if ($sync.Form -and $sync.Form.Dispatcher) {
        $sync.Form.Dispatcher.Invoke(
            [action]$Code,
            [System.Windows.Threading.DispatcherPriority]::Normal
        )
    }
}

function Set-WPFStatus {
    param([string]$Message)
    Invoke-WPFUIThread {
        if ($sync.WPFStatusBar) {
            $sync.WPFStatusBar.Text = $Message
        }
    }
}
