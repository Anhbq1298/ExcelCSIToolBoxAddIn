using System;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    internal static class BatchProgressWindow
    {
        public static BatchProgressResult RunWithProgress(int totalItems, string title, Action<BatchProgressContext> workAction)
        {
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

    internal sealed class BatchProgressContext
    {
        private int _ranCount;
        private int _skippedCount;
        private bool _cancellationRequested;

        public BatchProgressContext(int totalItems)
        {
            TotalItems = totalItems;
        }

        public bool IsCancellationRequested
        {
            get { return _cancellationRequested; }
        }

        public int RanCount
        {
            get { return _ranCount; }
        }

        public int SkippedCount
        {
            get { return _skippedCount; }
        }

        public int TotalItems { get; private set; }

        public void IncrementRan()
        {
            _ranCount++;
        }

        public void IncrementSkipped()
        {
            _skippedCount++;
        }

        public void RequestCancellation()
        {
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
