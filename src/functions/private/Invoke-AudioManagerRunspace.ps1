function Invoke-AudioManagerRunspace {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock)
    $ps = [powershell]::Create()
    $ps.RunspacePool = $sync.RunspacePool
    $ps.AddScript($ScriptBlock) | Out-Null
    $ps.BeginInvoke() | Out-Null
}
