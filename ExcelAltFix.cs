using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

internal static class NativeExcelAlt
{
    [STAThread]
    private static int Main()
    {
        bool created;
        using (var mutex = new Mutex(true, @"Local\NativeExcelAltRibbonFix", out created))
        {
            if (!created) return 0;
            try
            {
                try { SetProcessDpiAwarenessContext(new IntPtr(-4)); }
                catch (EntryPointNotFoundException) { SetProcessDPIAware(); }
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
                using (var context = new HelperContext()) Application.Run(context);
                return 0;
            }
            catch { return 1; }
        }
    }

    private sealed class HelperContext : ApplicationContext
    {
        private const uint InjectionTag = 0xCE1A17F0;
        private const int LAlt = 0xA4;
        private readonly HookProc keyboardProc;
        private readonly HookProc mouseProc;
        private IntPtr keyboardHook;
        private IntPtr mouseHook;
        private NotifyIcon tray;
        private bool enabled = true;
        private bool pending;
        private bool forwardedAlt;
        private IntPtr origin;
        private KeyboardEvent bufferedAlt;

        internal HelperContext()
        {
            keyboardProc = KeyboardHook;
            mouseProc = MouseHook;
            try
            {
                IntPtr module = GetModuleHandle(null);
                keyboardHook = SetWindowsHookEx(13, keyboardProc, module, 0);
                if (keyboardHook == IntPtr.Zero) throw new Win32Exception();
                mouseHook = SetWindowsHookEx(14, mouseProc, module, 0);
                if (mouseHook == IntPtr.Zero) throw new Win32Exception();
                var menu = new ContextMenuStrip();
                menu.Items.Add("Exit", null, delegate { ExitThread(); });
                tray = new NotifyIcon { Icon = SystemIcons.Application,
                    Text = "Excel Alt shortcut helper", ContextMenuStrip = menu, Visible = true };
            }
            catch { Dispose(); throw; }
        }

        private IntPtr KeyboardHook(int code, IntPtr message, IntPtr data)
        {
            if (code < 0) return Next(keyboardHook, code, message, data);
            KeyboardEvent key = (KeyboardEvent)Marshal.PtrToStructure(data, typeof(KeyboardEvent));
            bool up = message.ToInt64() == 0x101 || message.ToInt64() == 0x105;
            bool leftAlt = key.vk == LAlt || (key.vk == 0x12 && (key.flags & 1) == 0);
            if (key.extraInfo.ToUInt64() == InjectionTag)
            {
                return Next(keyboardHook, code, message, data);
            }
            try
            {
                if (forwardedAlt)
                {
                    if (leftAlt && up) forwardedAlt = false;
                    return Next(keyboardHook, code, message, data);
                }
                if (pending)
                {
                    if (leftAlt && !up) return new IntPtr(1);
                    if (leftAlt)
                    {
                        pending = false;
                        if (enabled && GetForegroundWindow() == origin)
                        {
                            Input down = KeyInput(0x79, 0x44, 0, false);
                            Input release = KeyInput(0x79, 0x44, 0, true);
                            uint sent = Send(new[] { down, release });
                            if (sent != 2)
                            {
                                if (sent == 1) Send(new[] { release });
                                Disable();
                            }
                        }
                        return new IntPtr(1);
                    }
                    pending = false;
                    uint count = Send(new[] { KeyInput(bufferedAlt, false), KeyInput(key, up) });
                    forwardedAlt = count > 0;
                    if (count == 2) return new IntPtr(1);
                    if (count == 0) forwardedAlt = Send(new[] { KeyInput(bufferedAlt, false) }) == 1;
                    Disable();
                    return Next(keyboardHook, code, message, data);
                }
                if (enabled && leftAlt && !up && !OtherInputHeld() && !Down(LAlt))
                {
                    IntPtr foreground = GetForegroundWindow();
                    if (IsExcel(foreground))
                    {
                        origin = foreground;
                        bufferedAlt = key;
                        pending = true;
                        return new IntPtr(1);
                    }
                }
            }
            catch
            {
                RestorePendingAlt();
                Disable();
            }
            return Next(keyboardHook, code, message, data);
        }

        private IntPtr MouseHook(int code, IntPtr message, IntPtr data)
        {
            if (code < 0 || !pending) return Next(mouseHook, code, message, data);
            MouseEvent mouse = (MouseEvent)Marshal.PtrToStructure(data, typeof(MouseEvent));
            if (mouse.extraInfo.ToUInt64() == InjectionTag) return Next(mouseHook, code, message, data);
            uint flags = MouseFlags(message.ToInt32());
            if (flags == 0) return Next(mouseHook, code, message, data);
            try
            {
                Input replay = MouseInput(mouse, flags);
                pending = false;
                uint count = Send(new[] { KeyInput(bufferedAlt, false), replay });
                forwardedAlt = count > 0;
                if (count == 2) return new IntPtr(1);
                if (count == 0) forwardedAlt = Send(new[] { KeyInput(bufferedAlt, false) }) == 1;
                Disable();
            }
            catch
            {
                RestorePendingAlt();
                Disable();
            }
            return Next(mouseHook, code, message, data);
        }

        private void RestorePendingAlt()
        {
            if (!pending) return;
            pending = false;
            forwardedAlt = Send(new[] { KeyInput(bufferedAlt, false) }) == 1;
        }

        private void Disable()
        {
            enabled = false;
            pending = false;
            if (tray != null) tray.Text = "Excel Alt helper disabled after input failure";
        }

        protected override void Dispose(bool disposing)
        {
            if (keyboardHook != IntPtr.Zero) { UnhookWindowsHookEx(keyboardHook); keyboardHook = IntPtr.Zero; }
            if (mouseHook != IntPtr.Zero) { UnhookWindowsHookEx(mouseHook); mouseHook = IntPtr.Zero; }
            if (forwardedAlt)
            {
                Send(new[] { KeyInput(LAlt, 0x38, 0, true) });
                forwardedAlt = false;
            }
            pending = false;
            if (tray != null)
            {
                tray.Visible = false;
                ContextMenuStrip menu = tray.ContextMenuStrip;
                tray.Dispose();
                menu.Dispose();
                tray = null;
            }
            base.Dispose(disposing);
        }


        private static bool IsExcel(IntPtr window)
        {
            if (window == IntPtr.Zero) return false;
            var name = new StringBuilder(64);
            if (GetClassName(window, name, name.Capacity) == 0 || name.ToString() != "XLMAIN") return false;
            uint pid;
            GetWindowThreadProcessId(window, out pid);
            if (!IsUnelevated(pid)) return false;
            using (Process process = Process.GetProcessById((int)pid))
                return string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsUnelevated(uint pid)
        {
            IntPtr process = OpenProcess(0x1000, false, pid);
            if (process == IntPtr.Zero) return false;
            IntPtr token = IntPtr.Zero;
            try
            {
                if (!OpenProcessToken(process, 0x8, out token)) return false;
                int elevated, length;
                return GetTokenInformation(token, 20, out elevated, 4, out length) && elevated == 0;
            }
            finally
            {
                if (token != IntPtr.Zero) CloseHandle(token);
                CloseHandle(process);
            }
        }

        private static bool OtherInputHeld()
        {
            return Down(0x10) || Down(0x11) || Down(0xA5) || Down(0x5B) || Down(0x5C)
                || Down(1) || Down(2) || Down(4) || Down(5) || Down(6);
        }

        private static bool Down(int key) { return (GetAsyncKeyState(key) & 0x8000) != 0; }
        private static IntPtr Next(IntPtr hook, int code, IntPtr message, IntPtr data)
        { return CallNextHookEx(hook, code, message, data); }
        private static uint Send(Input[] inputs)
        { return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(Input))); }
        private static Input KeyInput(KeyboardEvent key, bool up)
        { return KeyInput(key.vk, key.scan, key.flags & 1, up); }
        private static Input KeyInput(uint vk, uint scan, uint extended, bool up)
        {
            return new Input { type = 1, data = new InputData { keyboard = new KeyboardInput {
                vk = (ushort)vk, scan = (ushort)scan, flags = extended | (up ? 2u : 0u),
                extraInfo = new UIntPtr(InjectionTag) } } };
        }

        private static uint MouseFlags(int message)
        {
            switch (message)
            {
                case 0x201: return 0x2;
                case 0x202: return 0x4;
                case 0x204: return 0x8;
                case 0x205: return 0x10;
                case 0x207: return 0x20;
                case 0x208: return 0x40;
                case 0x20A: return 0x800;
                case 0x20B: return 0x80;
                case 0x20C: return 0x100;
                case 0x20E: return 0x1000;
                default: return 0;
            }
        }

        private static Input MouseInput(MouseEvent mouse, uint flags)
        {
            int left = GetSystemMetrics(76), top = GetSystemMetrics(77);
            int width = GetSystemMetrics(78), height = GetSystemMetrics(79);
            int x = (int)Math.Round((mouse.point.x - left) * 65535.0 / Math.Max(1, width - 1));
            int y = (int)Math.Round((mouse.point.y - top) * 65535.0 / Math.Max(1, height - 1));
            uint mouseData = (flags & 0x1800) != 0
                ? unchecked((uint)(short)(mouse.mouseData >> 16)) : mouse.mouseData >> 16;
            return new Input { type = 0, data = new InputData { mouse = new MouseInput {
                x = Math.Max(0, Math.Min(65535, x)), y = Math.Max(0, Math.Min(65535, y)),
                mouseData = mouseData, flags = flags | 0xC001,
                extraInfo = new UIntPtr(InjectionTag) } } };
        }
    }

    private delegate IntPtr HookProc(int code, IntPtr message, IntPtr data);
    [StructLayout(LayoutKind.Sequential)] private struct Point { internal int x, y; }
    [StructLayout(LayoutKind.Sequential)] private struct KeyboardEvent
    { internal uint vk, scan, flags, time; internal UIntPtr extraInfo; }
    [StructLayout(LayoutKind.Sequential)] private struct MouseEvent
    { internal Point point; internal uint mouseData, flags, time; internal UIntPtr extraInfo; }
    [StructLayout(LayoutKind.Sequential)] private struct Input
    { internal uint type; internal InputData data; }
    [StructLayout(LayoutKind.Explicit)] private struct InputData
    { [FieldOffset(0)] internal MouseInput mouse; [FieldOffset(0)] internal KeyboardInput keyboard; }
    [StructLayout(LayoutKind.Sequential)] private struct MouseInput
    { internal int x, y; internal uint mouseData, flags, time; internal UIntPtr extraInfo; }
    [StructLayout(LayoutKind.Sequential)] private struct KeyboardInput
    { internal ushort vk, scan; internal uint flags, time; internal UIntPtr extraInfo; }
    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int id, HookProc proc, IntPtr module, uint thread);
    [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
    [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint count, Input[] inputs, int size);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll")] private static extern bool SetProcessDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string name);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool GetTokenInformation(IntPtr token, int informationClass, out int information, int size, out int length);
}
