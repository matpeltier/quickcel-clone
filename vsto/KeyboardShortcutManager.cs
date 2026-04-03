using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;

namespace QuickCelVSTO
{
    public class KeyboardShortcutManager : IDisposable
    {
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc _proc;
        private IntPtr _hookId = IntPtr.Zero;

        private readonly Dictionary<int, Action> _shortcutMap = new Dictionary<int, Action>();
        private readonly HashSet<int> _registeredCombos = new HashSet<int>();

        private bool _ctrlPressed;
        private bool _shiftPressed;
        private bool _altPressed;

        public KeyboardShortcutManager()
        {
            _proc = HookCallback;
        }

        public void RegisterShortcut(bool ctrl, bool shift, bool alt, Keys key, Action action)
        {
            int comboHash = MakeHash(ctrl, shift, alt, key);
            _shortcutMap[comboHash] = action;
            _registeredCombos.Add(comboHash);
        }

        public void Start()
        {
            if (_hookId != IntPtr.Zero) return;

            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                _hookId = SetWindowsHookEx(
                    WH_KEYBOARD_LL,
                    _proc,
                    GetModuleHandle(curModule.ModuleName),
                    0);
            }
        }

        public void Stop()
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;

                if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey)
                {
                    _ctrlPressed = true;
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);
                }
                if (key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey)
                {
                    _shiftPressed = true;
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);
                }
                if (key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu)
                {
                    _altPressed = true;
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);
                }

                int hash = MakeHash(_ctrlPressed, _shiftPressed, _altPressed, NormalizeKey(key));
                if (_registeredCombos.Contains(hash))
                {
                    Action action;
                    if (_shortcutMap.TryGetValue(hash, out action))
                    {
                        try
                        {
                            action();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"QuickCel shortcut error: {ex.Message}");
                        }
                        return (IntPtr)1;
                    }
                }
            }
            else if (nCode >= 0 && (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;

                if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey)
                    _ctrlPressed = false;
                if (key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey)
                    _shiftPressed = false;
                if (key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu)
                    _altPressed = false;
            }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private Keys NormalizeKey(Keys key)
        {
            Keys k = key & Keys.KeyCode;
            if (k >= Keys.D0 && k <= Keys.D9) return k;
            if (k >= Keys.A && k <= Keys.Z) return k;
            if (k == Keys.Oemtilde) return Keys.Oemtilde;
            if (k == Keys.Oemcomma) return Keys.Oemcomma;
            if (k == Keys.OemPeriod) return Keys.OemPeriod;
            if (k == Keys.OemOpenBrackets) return Keys.OemOpenBrackets;
            if (k == Keys.OemCloseBrackets) return Keys.OemCloseBrackets;
            if (k == Keys.Down) return Keys.Down;
            if (k == Keys.Left) return Keys.Left;
            if (k == Keys.Right) return Keys.Right;
            if (k == Keys.Up) return Keys.Up;
            return k;
        }

        private static int MakeHash(bool ctrl, bool shift, bool alt, Keys key)
        {
            int h = (int)key;
            if (ctrl) h |= 0x10000;
            if (shift) h |= 0x20000;
            if (alt) h |= 0x40000;
            return h;
        }

        public void Dispose()
        {
            Stop();
        }

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}
