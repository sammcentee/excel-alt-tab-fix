$ErrorActionPreference = 'Stop'
$targetExe = Join-Path $env:USERPROFILE '.local\share\ExcelAltFix\ExcelAltFix.exe'
$startupLink = Join-Path ([Environment]::GetFolderPath('Startup')) 'Excel Alt shortcut fix.lnk'
if (Test-Path -LiteralPath $startupLink) {
    $shell = New-Object -ComObject WScript.Shell
    if ($shell.CreateShortcut($startupLink).TargetPath -ne $targetExe) {
        throw 'The startup shortcut does not point to this helper.'
    }
    Remove-Item -LiteralPath $startupLink
}
Get-CimInstance Win32_Process -Filter "Name = 'ExcelAltFix.exe'" |
    Where-Object { $_.ExecutablePath -eq $targetExe } |
    ForEach-Object { Stop-Process -Id $_.ProcessId }
'Excel Alt fix is disabled. The source files remain available.'

