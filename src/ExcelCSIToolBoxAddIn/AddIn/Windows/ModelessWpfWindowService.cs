using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms.Integration;
using System.Windows.Interop;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class ModelessWpfWindowService
    {
        private static readonly HashSet<Window> ConfiguredWindows = new HashSet<Window>();
        private static readonly object ConfigurationLock = new object();

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

                AssignExcelOwner(window);
                ElementHost.EnableModelessKeyboardInterop(window);
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
