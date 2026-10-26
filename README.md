# Excel Alt+Tab fix

Excel can ignore the first Alt tap after you use Alt+Tab to return to it.
This helper makes **one left-Alt tap** show the ribbon KeyTips (shortcut letters) again.

It sends F10 when you tap left Alt by itself in Excel. Alt combinations pass through.
You can also press [F10 directly](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel) without this helper.

## Install

1. [Download the ZIP](https://github.com/sammcentee/excel-alt-tab-fix/archive/refs/heads/main.zip) and extract it.
2. Open the folder that contains `Install.ps1`.
3. Right-click an empty area of the folder. Select **Open in Terminal**.
4. Run:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1
   ```

The script compiles the C# source with the .NET Framework compiler that Windows includes. No administrator rights or Office add-in required.
It installs to `%USERPROFILE%\.local\share\ExcelAltFix`, starts the helper, and adds a Startup shortcut for future sign-ins.

**To update:** download the new files and run the installer again.

## Use

Return to Excel with Alt+Tab, then tap left Alt. Tap it again to hide the KeyTips.
The helper runs in the background. It does not record, save, or transmit keystrokes.

## Stop or remove

For a temporary stop, right-click the helper's icon in the notification area and select **Exit**.
It starts again at your next sign-in.

To stop it and remove the Startup shortcut, run this command in PowerShell:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.local\share\ExcelAltFix\Disable.ps1"
```

The files stay in the install folder. To enable the helper again, run `Install.ps1` from that folder.
For full removal, delete the install folder after you run `Disable.ps1`.

## Compatibility and limits

- Requires Windows x64, .NET Framework 4.x, and desktop Excel. Tests used Windows 11 Home 25H2 and Microsoft 365 Excel x64.
- Only normal Excel main windows receive the fix. It excludes Excel dialogs, Excel that runs as administrator, and other apps.
- Right Alt and AltGr keep their normal behavior.
- Tests used scripted input on one PC. Physical keyboard timing, Alt+mouse, Alt+Shift+Tab, Remote Desktop, and other setups remain untested.
- Company policies or security software can block the scripts or helper. `-ExecutionPolicy Bypass` does not override Group Policy.

If the tooltip shows `Excel Alt helper disabled after input failure`, run `Install.ps1` again to restart the helper.
For other problems, [open an issue](https://github.com/sammcentee/excel-alt-tab-fix/issues) with your Windows version, Excel version, and steps to reproduce.

## License

[MIT](LICENSE). This project is not affiliated with Microsoft.
