using System;

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
        public static Func<int, string, Action<IBatchProgressContext>, BatchProgressSummary> ProgressRunner { get; set; }

        internal static BatchProgressResult RunWithProgress(int totalItems, string title, Action<BatchProgressContext> workAction)
        {
            if (ProgressRunner != null)
            {
                BatchProgressSummary summary = ProgressRunner(
                    totalItems,
                    title,
                    progressContext => workAction(new BatchProgressContext(totalItems, progressContext)));

                return new BatchProgressResult
                {
                    RanCount = summary == null ? 0 : summary.RanCount,
                    SkippedCount = summary == null ? 0 : summary.SkippedCount,
                    WasCancelled = summary != null && summary.WasCancelled
                };
            }

            BatchProgressContext context = new BatchProgressContext(totalItems);
            workAction(context);

            return new BatchProgressResult
            {
                RanCount = context.RanCount,
                SkippedCount = context.SkippedCount,
                WasCancelled = context.IsCancellationRequested
            };
        }
    }

    internal static class BatchProgressWindow
    {
        public static BatchProgressResult RunWithProgress(int totalItems, string title, Action<BatchProgressContext> workAction)
        {
            return BatchProgressHost.RunWithProgress(totalItems, title, workAction);
        }
    }

    internal sealed class BatchProgressContext : IBatchProgressContext
    {
        private readonly IBatchProgressContext _progressContext;
        private int _ranCount;
        private int _skippedCount;
        private bool _cancellationRequested;

        public BatchProgressContext(int totalItems)
        {
            TotalItems = totalItems;
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
        }

        public void IncrementSkipped()
        {
            if (_progressContext != null)
            {
                _progressContext.IncrementSkipped();
                return;
            }

            _skippedCount++;
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
