using System;
using System.Windows;
using System.Windows.Threading;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    /// <summary>
    /// A reusable batch progress dialog that shows a progress bar, percentage,
    /// and Ran/Skipped counters while executing a looping operation.
    ///
    /// Because ETABS COM calls must stay on the UI (STA) thread, this dialog
    /// executes work synchronously on the calling thread and pumps the WPF
    /// dispatcher after each progress update so the UI remains responsive.
    ///
    /// Usage:
    ///   var result = BatchProgressWindow.RunWithProgress(items.Count, (ctx) =>
    ///   {
    ///       foreach (var item in items)
    ///       {
    ///           if (ctx.IsCancellationRequested) break;
    ///
    ///           bool success = DoWork(item);
    ///           if (success)
    ///               ctx.IncrementRan();
    ///           else
    ///               ctx.IncrementSkipped();
    ///       }
    ///   });
    ///
    ///   // result.WasCancelled, result.RanCount, result.SkippedCount
    /// </summary>
    public partial class BatchProgressWindow : Window
    {
        private BatchProgressContext _context;

        private BatchProgressWindow(int totalItems)
        {
            InitializeComponent();
            MainProgressBar.Maximum = totalItems;
            MainProgressBar.Value = 0;
            PercentageText.Text = $"0% (0/{totalItems})";
            RanCountRun.Text = "0";
            SkippedCountRun.Text = "0";
        }

        /// <summary>
        /// Runs a batch operation with a progress dialog.
        /// The <paramref name="workAction"/> is executed synchronously on the UI thread.
        /// Call methods on the provided <see cref="BatchProgressContext"/> to report progress.
        /// The dialog pumps the WPF dispatcher after each update so the UI stays responsive
        /// and the Cancel button remains clickable.
        /// </summary>
        /// <param name="totalItems">Total number of items to process (for percentage calculation).</param>
        /// <param name="workAction">
        /// The work to perform. Receives a <see cref="BatchProgressContext"/> that must be used
        /// to report Ran/Skipped counts and to check for cancellation.
        /// </param>
        /// <returns>A <see cref="BatchProgressResult"/> with the final Ran, Skipped counts and cancellation status.</returns>
        public static BatchProgressResult RunWithProgress(int totalItems, Action<BatchProgressContext> workAction)
        {
            var window = new BatchProgressWindow(totalItems);
            var context = new BatchProgressContext(totalItems, window);
            window._context = context;

            // Show non-modal so the UI thread can continue executing workAction below.
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

    /// <summary>
    /// Context object passed to the batch work action.
    /// Provides methods to report progress and check cancellation.
    /// After each increment call, the WPF dispatcher is pumped so the UI
    /// repaints and the Cancel button can be clicked.
    /// </summary>
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

        /// <summary>
        /// Returns true if the user has clicked Cancel.
        /// Check this at the start of each loop iteration.
        /// </summary>
        public bool IsCancellationRequested => _cancellationRequested;

        /// <summary>Current count of successfully processed items.</summary>
        public int RanCount => _ranCount;

        /// <summary>Current count of skipped items.</summary>
        public int SkippedCount => _skippedCount;

        /// <summary>Total items to process.</summary>
        public int TotalItems => _totalItems;

        /// <summary>
        /// Call this after successfully processing an item.
        /// Increments the Ran counter and updates the progress bar.
        /// </summary>
        public void IncrementRan()
        {
            _ranCount++;
            PumpUI();
        }

        /// <summary>
        /// Call this when an item is skipped (e.g. already exists).
        /// Increments the Skipped counter and updates the progress bar.
        /// </summary>
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
            // Update the progress window controls
            _window.UpdateProgressUI(_ranCount, _skippedCount);

            // Pump the WPF dispatcher so the window repaints and processes
            // input events (e.g. the Cancel button click).
            _window.Dispatcher.Invoke(DispatcherPriority.Background, new Action(delegate { }));
        }
    }

    /// <summary>
    /// Result returned after the batch progress dialog completes.
    /// </summary>
    public class BatchProgressResult
    {
        /// <summary>Number of items that were successfully processed.</summary>
        public int RanCount { get; set; }

        /// <summary>Number of items that were skipped.</summary>
        public int SkippedCount { get; set; }

        /// <summary>True if the user cancelled the operation before it completed.</summary>
        public bool WasCancelled { get; set; }

        /// <summary>Total processed items (Ran + Skipped).</summary>
        public int TotalProcessed => RanCount + SkippedCount;
    }
}
