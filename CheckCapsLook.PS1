Add-Type -AssemblyName System.Windows.Forms

if ([System.Windows.Forms.Control]::IsKeyLocked('CapsLock')) {
    # Simula pressionar Caps Lock para desativar
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys('{CAPSLOCK}')
    Write-Output "Caps Lock desativado"
} else {
    Write-Output "Caps Lock jÃ¡ estava desativado"
}
