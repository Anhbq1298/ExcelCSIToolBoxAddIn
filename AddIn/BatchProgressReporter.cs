using System;
using System.Windows;
using System.Windows.Controls;
using ExcelCSIToolBox.Core.Abstractions;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    /// <summary>
    /// WPF progress reporter used by infrastructure batch operations.
    /// </summary>
    internal sealed class BatchProgressReporter : IProgressReporter
    {
        private Window _window;
        private ProgressBar _progressBar;
        private TextBlock _messageText;

        public void Report(int percent, string message)
        {
            InvokeOnDispatcher(() =>
            {
                EnsureWindow();
                _progressBar.Value = Math.Max(0, Math.Min(100, percent));
                _messageText.Text = message ?? string.Empty;
            });
        }

        public void ReportComplete(string message)
        {
            InvokeOnDispatcher(() =>
            {
                if (_window == null)
                {
                    return;
                }

                _progressBar.Value = 100;
                _messageText.Text = message ?? "Completed.";
                _window.Close();
                _window = null;
            });
        }

        public void ReportError(string message)
        {
            InvokeOnDispatcher(() =>
            {
                if (_window == null)
                {
                    EnsureWindow();
                }

                _messageText.Text = message ?? "Operation failed.";
                _window.Close();
                _window = null;
            });
        }

        private void EnsureWindow()
        {
            if (_window != null)
            {
                return;
            }

            _messageText = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 10),
                TextWrapping = TextWrapping.Wrap
            };

            _progressBar = new ProgressBar
            {
                Minimum = 0,
                Maximum = 100,
                Height = 18
            };

            _window = new Window
            {
                Title = "Batch Progress",
                Width = 420,
                Height = 120,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Content = new StackPanel
                {
                    Margin = new Thickness(16),
                    Children =
                    {
                        _messageText,
                        _progressBar
                    }
                }
            };

            _window.Show();
        }

        private static void InvokeOnDispatcher(Action action)
        {
            new WpfThreadDispatcher().InvokeOnUiThread(action);
        }
    }
}
