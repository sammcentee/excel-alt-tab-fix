$ErrorActionPreference = 'Stop'
$installRoot = Join-Path $env:USERPROFILE '.local\share\ExcelAltFix'
$targetExe = Join-Path $installRoot 'ExcelAltFix.exe'
$startupLink = Join-Path ([Environment]::GetFolderPath('Startup')) 'Excel Alt shortcut fix.lnk'
$shell = New-Object -ComObject WScript.Shell
if (Test-Path -LiteralPath $startupLink) {
    if ($shell.CreateShortcut($startupLink).TargetPath -ne $targetExe) {
        throw 'The startup name belongs to another program.'
    }
}
New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
foreach ($name in @('ExcelAltFix.cs', 'Install.ps1', 'Disable.ps1')) {
    $source = Join-Path $PSScriptRoot $name
    $target = Join-Path $installRoot $name
    if ([IO.Path]::GetFullPath($source) -ne [IO.Path]::GetFullPath($target)) {
        Copy-Item -LiteralPath $source -Destination $target -Force
    }
}
Get-CimInstance Win32_Process -Filter "Name = 'ExcelAltFix.exe'" |
    Where-Object { $_.ExecutablePath -eq $targetExe } |
    ForEach-Object { Stop-Process -Id $_.ProcessId }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /warnaserror /platform:x64 /target:winexe /reference:System.Windows.Forms.dll /reference:System.Drawing.dll "/out:$targetExe" (Join-Path $installRoot 'ExcelAltFix.cs')
if ($LASTEXITCODE -ne 0) { throw 'The helper did not compile. Startup was not changed.' }
$shortcut = $shell.CreateShortcut($startupLink)
$shortcut.TargetPath = $targetExe
$shortcut.WorkingDirectory = $installRoot
$shortcut.Description = 'Use one left-Alt tap for Excel ribbon shortcuts.'
$shortcut.WindowStyle = 7
$shortcut.Save()
Start-Process -FilePath $targetExe -WindowStyle Hidden
'Excel Alt fix is enabled now and at sign-in.'

