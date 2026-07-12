using System;
using ExcelCSIToolBox.Core.Abstractions;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public interface IBatchProgressContext
    {
        bool IsCancellationRequested { get; }
        int RanCount { get; }
        int SkippedCount { get; }
        int TotalItems { get; }
        void IncrementRan();
        void IncrementSkipped();
        void RequestCancellation();
    }

    public sealed class BatchProgressSummary
    {
        public int RanCount { get; set; }
        public int SkippedCount { get; set; }
        public bool WasCancelled { get; set; }
    }

    public static class BatchProgressHost
    {
        // Progress reporting migrated to injected IProgressReporter.
    }

    internal static class BatchProgressWindow
    {
        public static BatchProgressResult RunWithProgress(
            int totalItems,
            string title,
            Action<BatchProgressContext> workAction,
            IProgressReporter progressReporter = null)
        {
            BatchProgressContext context = new BatchProgressContext(totalItems, title, progressReporter);
            try
            {
                workAction(context);
                progressReporter?.ReportComplete($"{title} completed.");
            }
            catch (Exception ex)
            {
                progressReporter?.ReportError($"{title} failed: {ex.Message}");
                throw;
            }

            return new BatchProgressResult
            {
                RanCount = context.RanCount,
                SkippedCount = context.SkippedCount,
                WasCancelled = context.IsCancellationRequested
            };
        }
    }

    internal sealed class BatchProgressContext : IBatchProgressContext
    {
        private readonly IBatchProgressContext _progressContext;
        private readonly IProgressReporter _progressReporter;
        private readonly string _title;
        private int _ranCount;
        private int _skippedCount;
        private bool _cancellationRequested;

        public BatchProgressContext(int totalItems)
        {
            TotalItems = totalItems;
        }

        public BatchProgressContext(int totalItems, string title, IProgressReporter progressReporter)
        {
            TotalItems = totalItems;
            _title = title;
            _progressReporter = progressReporter;
            ReportProgress();
        }

        internal BatchProgressContext(int totalItems, IBatchProgressContext progressContext)
        {
            TotalItems = totalItems;
            _progressContext = progressContext;
        }

        public bool IsCancellationRequested
        {
            get { return _progressContext == null ? _cancellationRequested : _progressContext.IsCancellationRequested; }
        }

        public int RanCount
        {
            get { return _progressContext == null ? _ranCount : _progressContext.RanCount; }
        }

        public int SkippedCount
        {
            get { return _progressContext == null ? _skippedCount : _progressContext.SkippedCount; }
        }

        public int TotalItems { get; private set; }

        public void IncrementRan()
        {
            if (_progressContext != null)
            {
                _progressContext.IncrementRan();
                return;
            }

            _ranCount++;
            ReportProgress();
        }

        public void IncrementSkipped()
        {
            if (_progressContext != null)
            {
                _progressContext.IncrementSkipped();
                return;
            }

            _skippedCount++;
            ReportProgress();
        }

        public void RequestCancellation()
        {
            if (_progressContext != null)
            {
                _progressContext.RequestCancellation();
                return;
            }

            _cancellationRequested = true;
        }

        private void ReportProgress()
        {
            if (_progressReporter == null)
            {
                return;
            }

            int processed = RanCount + SkippedCount;
            int percent = TotalItems <= 0 ? 0 : (int)Math.Round(100.0 * processed / TotalItems);
            _progressReporter.Report(percent, $"{_title} ({processed}/{TotalItems})");
        }
    }

    internal sealed class BatchProgressResult
    {
        public int RanCount { get; set; }
        public int SkippedCount { get; set; }
        public bool WasCancelled { get; set; }

        public int TotalProcessed
        {
            get { return RanCount + SkippedCount; }
        }
    }
}
