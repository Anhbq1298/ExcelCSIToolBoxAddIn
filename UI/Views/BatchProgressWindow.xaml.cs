using System;
using System.Windows;
using System.Windows.Threading;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// A reusable batch progress dialog that shows a progress bar, percentage,
    /// and Ran/Skipped counters while executing a looping operation.
    /// </summary>
    public partial class BatchProgressWindow : Window
    {
        private BatchProgressContext _context;

        private BatchProgressWindow(int totalItems, string title)
        {
            InitializeComponent();
            TaskTitleText.Text = title;
            MainProgressBar.Maximum = totalItems;
            MainProgressBar.Value = 0;
            PercentageText.Text = $"0% (0/{totalItems})";
            RanCountRun.Text = "0";
            SkippedCountRun.Text = "0";
        }

        /// <summary>
        /// Runs a batch operation with a progress dialog.
        /// </summary>
        /// <param name="totalItems">Total number of items to process.</param>
        /// <param name="title">The title/description of the task (e.g. "Creating Steel Sections...").</param>
        /// <param name="workAction">The work to perform.</param>
        /// <returns>A BatchProgressResult.</returns>
        public static BatchProgressResult RunWithProgress(int totalItems, string title, Action<BatchProgressContext> workAction)
        {
            var window = new BatchProgressWindow(totalItems, title);
            var context = new BatchProgressContext(totalItems, window);
            window._context = context;

            window.Show();

            try
            {
                workAction(context);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BatchProgressWindow work action failed: {ex.Message}");
            }

            var result = new BatchProgressResult
            {
                RanCount = context.RanCount,
                SkippedCount = context.SkippedCount,
                WasCancelled = context.IsCancellationRequested
            };

            window.Close();
            return result;
        }

        internal void UpdateProgressUI(int ranCount, int skippedCount)
        {
            int processed = ranCount + skippedCount;
            int total = (int)MainProgressBar.Maximum;

            MainProgressBar.Value = processed;
            int percent = total > 0 ? (int)Math.Round(100.0 * processed / total) : 0;
            PercentageText.Text = $"{percent}% ({processed}/{total})";
            RanCountRun.Text = ranCount.ToString();
            SkippedCountRun.Text = skippedCount.ToString();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context != null)
            {
                _context.RequestCancellation();
            }
            CancelButton.IsEnabled = false;
            CancelButton.Content = "Cancelling...";
        }
    }

    public class BatchProgressContext
    {
        private readonly int _totalItems;
        private readonly BatchProgressWindow _window;
        private int _ranCount;
        private int _skippedCount;
        private volatile bool _cancellationRequested;

        internal BatchProgressContext(int totalItems, BatchProgressWindow window)
        {
            _totalItems = totalItems;
            _window = window;
        }

        public bool IsCancellationRequested => _cancellationRequested;
        public int RanCount => _ranCount;
        public int SkippedCount => _skippedCount;
        public int TotalItems => _totalItems;

        public void IncrementRan()
        {
            _ranCount++;
            PumpUI();
        }

        public void IncrementSkipped()
        {
            _skippedCount++;
            PumpUI();
        }

        internal void RequestCancellation()
        {
            _cancellationRequested = true;
        }

        private void PumpUI()
        {
            _window.UpdateProgressUI(_ranCount, _skippedCount);
            _window.Dispatcher.Invoke(DispatcherPriority.Background, new Action(delegate { }));
        }
    }

    public class BatchProgressResult
    {
        public int RanCount { get; set; }
        public int SkippedCount { get; set; }
        public bool WasCancelled { get; set; }
        public int TotalProcessed => RanCount + SkippedCount;
    }
}

