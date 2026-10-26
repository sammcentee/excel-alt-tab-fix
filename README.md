# Excel Alt+Tab fix

This helper fixes an Excel keyboard problem on Windows.
After you use Alt+Tab to go back to Excel, Excel ignores the first Alt tap.
You must tap Alt two times before the KeyTips (the shortcut letters on the ribbon) come on.

With this helper, one tap of the left Alt key shows the KeyTips again.

ExcelAltFix is a small background helper: one C# file of about 300 lines.
`Install.ps1` compiles it on your PC with the C# compiler of the .NET Framework, which is part of Windows.
The helper needs no administrator rights, no Office add-in, no macro, and no registry change.

## The problem

1. Work in Excel.
2. Press Alt+Tab to go to a different window.
3. Press Alt+Tab again to go back to Excel.
4. Tap Alt.

Expected result: the KeyTips come on.
Actual result: nothing occurs. Only the second Alt tap shows the KeyTips.

Other users report the same problem on Microsoft Q&A: [Excel ignoring first keypress of ALT after task switching](https://learn.microsoft.com/en-us/answers/questions/4879388/excel-ignoring-first-keypress-of-alt-after-task-sw).
That thread also reports the problem in other Office apps. This helper changes only Excel.
We did not find a fix from Microsoft.

F10 does not have this problem.
Microsoft documents F10 as an alternative to Alt for the KeyTips ([Keyboard shortcuts in Excel](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)).
If you can use F10, you do not need this helper.
The helper is for people who want to keep the Alt habit.

## How the fix works

The helper monitors the keyboard and the mouse with Windows low-level input hooks.

- When you tap the left Alt key by itself in Excel, the helper stops that Alt tap and sends F10 to Excel. Excel then shows the KeyTips.
- When you hold the left Alt key and press a different key or a mouse button, the helper sends the held Alt key, and then that key or button. Thus Alt+Tab, Alt+H, Alt+Space, Alt+F4, and the other Alt shortcuts operate as usual.
- The helper acts only when the foreground window is an ordinary Excel main window (`EXCEL.EXE`, window class `XLMAIN`).
- The helper does not change the right Alt key or AltGr.
- The helper does not change other apps, Excel dialog boxes, or Excel that runs as administrator.
- If Shift, Ctrl, right Alt, the Windows key, or a mouse button is already down when you press left Alt, the helper does not change that Alt key.
- If Windows does not accept an input that the helper sends, or an unexpected error occurs, the helper stops changing the Alt key. Its tooltip then shows `Excel Alt helper disabled after input failure`. To enable it again, refer to [Troubleshooting](#troubleshooting).

The helper does not record, save, or transmit your keystrokes. It writes no files and uses no network connection.

## Requirements

- Windows 11 x64. The tests used Windows 11 Home 25H2. The helper possibly also operates on Windows 10 x64, but we did not do a test on Windows 10.
- Desktop Excel for Windows. The tests used Microsoft 365 Excel x64.
- .NET Framework 4.x, which is part of Windows 10 and Windows 11. `Install.ps1` uses the compiler `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe`.
- No administrator rights.

## Install

1. Get the files. Use one of these methods:
   - On this page, select **Code** > **Download ZIP**. Then extract the ZIP file.
   - Download the source code ZIP file from the [latest release](https://github.com/sammcentee/excel-alt-tab-fix/releases/latest). Then extract the ZIP file.
   - Clone the repository: `git clone https://github.com/sammcentee/excel-alt-tab-fix.git`
2. In File Explorer, open the folder that contains `Install.ps1`. A ZIP file puts the files in a subfolder, for example `excel-alt-tab-fix-main`.
3. Right-click an empty area of the folder, and then select **Open in Terminal**.
4. Run this command:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1
   ```

5. Make sure that the command shows `Excel Alt fix is enabled now and at sign-in.`

`Install.ps1` does these steps:

1. It copies `ExcelAltFix.cs`, `Install.ps1`, and `Disable.ps1` to the install folder, `%USERPROFILE%\.local\share\ExcelAltFix`.
2. It stops the helper if the helper runs from the install folder.
3. It compiles `ExcelAltFix.cs` to `ExcelAltFix.exe` in the install folder.
4. It adds the shortcut `Excel Alt shortcut fix.lnk` to your Startup folder. The helper then starts each time that you sign in.
5. It starts the helper.

If the compilation fails, the script stops. It does not change the Startup shortcut, and it does not start the helper. If the helper was in operation before, the script has already stopped it.

The repository contains no `.exe` file. You can read all of the code before you install it.

## Use

1. Press Alt+Tab to go back to Excel.
2. Tap the left Alt key one time. The KeyTips come on.
3. Tap the left Alt key again to remove the KeyTips.

The helper shows the standard Windows application icon in the notification area of the taskbar. Its tooltip is `Excel Alt shortcut helper`.
If you cannot see the icon, select the up arrow (**Show hidden icons**) on the taskbar.

## Update

1. Get the new files.
2. Do the [Install](#install) procedure again. `Install.ps1` replaces the files, compiles the helper again, and starts it.

## Stop or remove

- To stop the helper until the next sign-in, right-click its icon in the notification area, and then select **Exit**.
- To stop the helper and remove its Startup shortcut, run this command in PowerShell. The files stay in the install folder.

  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.local\share\ExcelAltFix\Disable.ps1"
  ```

- To enable the helper again, run this command in PowerShell:

  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.local\share\ExcelAltFix\Install.ps1"
  ```

- To remove the helper fully, run `Disable.ps1` first. Then delete the install folder, `%USERPROFILE%\.local\share\ExcelAltFix`.

## Troubleshooting

- If `Install.ps1` shows an error, copy the full error text into a new [issue](https://github.com/sammcentee/excel-alt-tab-fix/issues).
- To make sure that the helper is in operation, open Task Manager, select **Details**, and find `ExcelAltFix.exe`.
- If the tooltip shows `Excel Alt helper disabled after input failure`, start the helper again. Run the `Install.ps1` command in [Stop or remove](#stop-or-remove), or sign out of Windows and sign in again.
- If the helper does not start, your security software can possibly block it. Make sure that your security software permits `ExcelAltFix.exe`.
- If Excel runs as administrator, the helper does not change it. Start Excel without administrator rights.
- To report a different problem, open an [issue](https://github.com/sammcentee/excel-alt-tab-fix/issues). Give your Windows version, your Excel version (32-bit or 64-bit), and the steps that cause the problem.

## Test results

On 18 October 2026, we did these tests on one PC with Windows 11 Home 25H2 (build 26200.9711) and Microsoft 365 Excel x64 (version 16.0.20430.20182).
The tests used blank test workbooks and one test window.

- Excel to the test window and back: 10 of 10 tests passed. Each time, one Alt tap showed the KeyTips, and a second Alt tap removed them.
- Excel workbook to a different Excel workbook and back: 5 of 5 tests passed.
- These ribbon key sequences operated:
  - Alt, W, P (Page Layout view)
  - Alt, W, L (Normal view)
  - Alt, O, C, W (Column Width, an old Excel 2003 sequence)
- After a window change, Alt+H (hold Alt, then press H) opened the Home tab.
- Alt+Space opened the window menu. Alt+F4 closed the test workbooks, and it also closed the test window outside Excel.
- Alt then Shift, Alt then Ctrl, Ctrl then Alt, Shift then Alt, and right Alt gave the same result as Excel without the helper.
- After the modifier tests, none of the tested modifier keys stayed down.
- In cell edit mode (F2) and in the formula bar, the Alt result was the same as Excel without the helper.
- `Disable.ps1` removed the Startup shortcut and stopped the helper. `Install.ps1` added them again, and the next one-Alt test passed.

On 26 October 2026, we ran the published `Install.ps1` on the same PC, with the `RemoteSigned` execution policy:

- From a separate folder with files that had the internet download mark: the script installed the files, compiled the helper, added the Startup shortcut, and started the helper.
- From the install folder, with the command in [Stop or remove](#stop-or-remove): the script compiled and started the helper again.

This check did not test the Alt key again. The source code is the same as in the 18 October tests.

## Limits

This helper is a workaround. It is not a change to Excel. It must run in the background.

No set of tests can show that the helper has no conflicts with all apps, add-ins, and setups.
We did not do tests on these items:

- Keystrokes from a physical keyboard at different speeds. The tests used scripted input, with a check of the window state between steps.
- Mouse input together with Alt: a mouse click or drag while you hold Alt, or a mouse button that is down before you press Alt. We examined the code for these paths, but we did not do a test.
- Alt+Shift+Tab.
- Remote Desktop.
- More than one Excel process at the same time.
- Workbooks with content, and Excel add-ins.
- Other PCs, Windows 10, ARM PCs, 32-bit Excel, and other Excel versions.

On a PC that your company manages, Group Policy or application control can stop the scripts or the helper.
The `-ExecutionPolicy Bypass` option does not override an execution policy that Group Policy sets.
Some security software can block or report programs that use keyboard hooks.

## Files

| File | Purpose |
| --- | --- |
| `ExcelAltFix.cs` | The source code of the helper. |
| `Install.ps1` | Copies, compiles, and starts the helper, and adds the Startup shortcut. |
| `Disable.ps1` | Stops the helper and removes the Startup shortcut. It keeps the files. |

## License

MIT. Refer to [LICENSE](LICENSE).

This project is not related to Microsoft. Microsoft and Excel are trademarks of the Microsoft group of companies.
