using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public sealed class QuickCreatePileCapWindowService
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly Func<IntPtr> _ownerHandleProvider;
        private QuickCreatePileCapWindow _activeWindow;

        public QuickCreatePileCapWindowService(
            ICSISapModelConnectionService connectionService,
            Func<IntPtr> ownerHandleProvider)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException("connectionService");
            _ownerHandleProvider = ownerHandleProvider ?? delegate { return IntPtr.Zero; };
        }

        public void ShowWindow()
        {
            if (_activeWindow != null)
            {
                ActivateWindow(_activeWindow);
                return;
            }

            var viewModel = new QuickCreatePileCapViewModel(
                _connectionService,
                CloseActiveWindow,
                RestoreFocusToActiveWindow);
            var window = new QuickCreatePileCapWindow(viewModel)
            {
                ShowInTaskbar = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            AssignExcelOwner(window);
            window.Closed += OnWindowClosed;
            _activeWindow = window;

            window.ShowDialog();
        }

        public void RestoreFocusToActiveWindow()
        {
            QuickCreatePileCapWindow window = _activeWindow;
            if (window == null)
            {
                return;
            }

            if (!window.Dispatcher.CheckAccess())
            {
                window.Dispatcher.BeginInvoke(new Action(RestoreFocusToActiveWindow), DispatcherPriority.ApplicationIdle);
                return;
            }

            window.Dispatcher.BeginInvoke(new Action(delegate { ActivateWindow(window); }), DispatcherPriority.ApplicationIdle);
        }

        private void CloseActiveWindow()
        {
            QuickCreatePileCapWindow window = _activeWindow;
            if (window == null)
            {
                return;
            }

            if (!window.Dispatcher.CheckAccess())
            {
                window.Dispatcher.BeginInvoke(new Action(CloseActiveWindow), DispatcherPriority.Normal);
                return;
            }

            window.Close();
        }

        private void AssignExcelOwner(Window window)
        {
            IntPtr ownerHandle = IntPtr.Zero;
            try
            {
                ownerHandle = _ownerHandleProvider();
            }
            catch
            {
                ownerHandle = IntPtr.Zero;
            }

            if (ownerHandle == IntPtr.Zero)
            {
                ownerHandle = Process.GetCurrentProcess().MainWindowHandle;
            }

            var helper = new WindowInteropHelper(window);
            if (ownerHandle != IntPtr.Zero)
            {
                helper.Owner = ownerHandle;
            }

            helper.EnsureHandle();
        }

        private void ActivateWindow(QuickCreatePileCapWindow window)
        {
            if (window == null)
            {
                return;
            }

            if (window.WindowState == WindowState.Minimized)
            {
                window.WindowState = WindowState.Normal;
            }

            window.Activate();
            window.Focus();
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
            QuickCreatePileCapWindow window = sender as QuickCreatePileCapWindow;
            if (window != null)
            {
                window.Closed -= OnWindowClosed;
            }

            if (ReferenceEquals(_activeWindow, window))
            {
                _activeWindow = null;
            }
        }
    }
}
