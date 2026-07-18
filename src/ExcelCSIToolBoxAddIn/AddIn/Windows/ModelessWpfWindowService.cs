using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms.Integration;
using System.Windows.Interop;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class ModelessWpfWindowService
    {
        private static readonly HashSet<Window> ConfiguredWindows = new HashSet<Window>();
        private static readonly object ConfigurationLock = new object();

        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        public static void Show(Window window)
        {
            if (window == null)
            {
                throw new ArgumentNullException(nameof(window));
            }

            if (!window.Dispatcher.CheckAccess())
            {
                window.Dispatcher.Invoke(new Action(delegate { Show(window); }));
                return;
            }

            EnsureConfigured(window);

            if (window.WindowState == WindowState.Minimized)
            {
                window.WindowState = WindowState.Normal;
            }

            if (!window.IsVisible)
            {
                window.Show();
            }

            window.Activate();
            window.Focus();
        }

        public static void CloseAll()
        {
            Window[] windows;
            lock (ConfigurationLock)
            {
                windows = new Window[ConfiguredWindows.Count];
                ConfiguredWindows.CopyTo(windows);
            }

            foreach (Window window in windows)
            {
                try
                {
                    if (window.Dispatcher.CheckAccess())
                    {
                        window.Close();
                    }
                    else
                    {
                        window.Dispatcher.Invoke(new Action(window.Close));
                    }
                }
                catch (InvalidOperationException)
                {
                    // The window was already closing or its dispatcher has shut down.
                }
            }
        }

        private static HookProc _hookProc;
        private static IntPtr _hookHandle = IntPtr.Zero;

        [ThreadStatic]
        private static bool _isProcessingHook;

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, ref MSG lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, ref MSG lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern uint GetCurrentThreadId();

        private static void EnsureConfigured(Window window)
        {
            lock (ConfigurationLock)
            {
                if (ConfiguredWindows.Contains(window))
                {
                    return;
                }

                ElementHost.EnableModelessKeyboardInterop(window);
                AssignExcelOwner(window);

                if (ConfiguredWindows.Count == 0)
                {
                    _hookProc = GetMessageHookProc;
                    _hookHandle = SetWindowsHookEx(3, _hookProc, IntPtr.Zero, GetCurrentThreadId());
                }

                window.Closed += OnWindowClosed;
                ConfiguredWindows.Add(window);
            }
        }

        private static void OnWindowClosed(object sender, EventArgs e)
        {
            Window window = sender as Window;
            if (window == null)
            {
                return;
            }

            window.Closed -= OnWindowClosed;
            lock (ConfigurationLock)
            {
                ConfiguredWindows.Remove(window);

                if (ConfiguredWindows.Count == 0)
                {
                    if (_hookHandle != IntPtr.Zero)
                    {
                        UnhookWindowsHookEx(_hookHandle);
                        _hookHandle = IntPtr.Zero;
                        _hookProc = null;
                    }
                }
            }
        }

        private static IntPtr GetMessageHookProc(int code, IntPtr wParam, ref MSG msg)
        {
            if (code >= 0 && !_isProcessingHook)
            {
                _isProcessingHook = true;
                try
                {
                    // Keyboard messages (WM_KEYFIRST to WM_KEYLAST)
                    const int WM_KEYFIRST = 0x0100;
                    const int WM_KEYLAST = 0x0109;

                    if (msg.message >= WM_KEYFIRST && msg.message <= WM_KEYLAST)
                    {
                        // Safely check if target window belongs to current thread before executing heavy operations or IsChild
                        uint windowThreadId = GetWindowThreadProcessId(msg.hwnd, IntPtr.Zero);
                        if (windowThreadId == GetCurrentThreadId())
                        {
                            foreach (Window window in ConfiguredWindows)
                            {
                                try
                                {
                                    IntPtr windowHandle = new WindowInteropHelper(window).Handle;
                                    if (windowHandle != IntPtr.Zero)
                                    {
                                        bool isTargetedToWpf = msg.hwnd == windowHandle || IsChild(windowHandle, msg.hwnd);
                                        if (isTargetedToWpf)
                                        {
                                            bool handled = ComponentDispatcher.RaiseThreadMessage(ref msg);
                                            if (!handled)
                                            {
                                                TranslateMessage(ref msg);
                                                DispatchMessage(ref msg);
                                            }

                                            // Set to WM_NULL to discard from Excel's message loop
                                            msg.message = 0x0000;
                                            break;
                                        }
                                    }
                                }
                                catch
                                {
                                    // Safeguard against disposed states
                                }
                            }
                        }
                    }
                }
                finally
                {
                    _isProcessingHook = false;
                }
            }
            return CallNextHookEx(_hookHandle, code, wParam, ref msg);
        }

        private static void AssignExcelOwner(Window window)
        {
            try
            {
                if (Globals.ExcelCSIToolBoxAddin == null ||
                    Globals.ExcelCSIToolBoxAddin.Application == null)
                {
                    return;
                }

                IntPtr excelHandle = new IntPtr(Globals.ExcelCSIToolBoxAddin.Application.Hwnd);
                if (excelHandle != IntPtr.Zero)
                {
                    new WindowInteropHelper(window).Owner = excelHandle;
                }
            }
            catch
            {
                // Keyboard interop still works when Excel's native owner is unavailable.
            }
        }
    }
}
