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
                    ComponentDispatcher.ThreadFilterMessage += OnThreadFilterMessage;
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
                    ComponentDispatcher.ThreadFilterMessage -= OnThreadFilterMessage;
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage(ref MSG lpMsg);

        private static void OnThreadFilterMessage(ref MSG msg, ref bool handled)
        {
            if (handled)
            {
                return;
            }

            // Keyboard messages (WM_KEYFIRST to WM_KEYLAST)
            const int WM_KEYFIRST = 0x0100;
            const int WM_KEYLAST = 0x0109;

            if (msg.message >= WM_KEYFIRST && msg.message <= WM_KEYLAST)
            {
                lock (ConfigurationLock)
                {
                    foreach (Window window in ConfiguredWindows)
                    {
                        try
                        {
                            IntPtr windowHandle = new WindowInteropHelper(window).Handle;
                            if (windowHandle != IntPtr.Zero)
                            {
                                bool isTargetedToWpf = msg.hwnd == windowHandle || IsChild(windowHandle, msg.hwnd);
                                if (isTargetedToWpf || window.IsActive || window.IsKeyboardFocusWithin)
                                {
                                    // Call ComponentDispatcher.RaiseThreadMessage to let WPF controls process key events (Tab, Arrows, typing)
                                    handled = ComponentDispatcher.RaiseThreadMessage(ref msg);
                                    if (!handled && isTargetedToWpf)
                                    {
                                        TranslateMessage(ref msg);
                                        DispatchMessage(ref msg);
                                        handled = true;
                                    }
                                    if (handled)
                                    {
                                        return;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // Safeguard against disposed window states
                        }
                    }
                }
            }
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
