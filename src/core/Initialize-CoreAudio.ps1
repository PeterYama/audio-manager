try {
    Add-Type -TypeDefinition $script:CoreAudioCSharp -Language CSharp -ReferencedAssemblies @(
        'System',
        'System.Runtime.InteropServices'
    ) -ErrorAction Stop
} catch {
    Write-Warning "CoreAudio type already loaded or failed to load: $_"
}
