namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters
{
    using System.Collections.Generic;

    public interface ICsiModelAdapter
    {
        string ApplicationName { get; }

        IReadOnlyList<CsiAttachResult> GetRunningInstances();

        CsiAttachResult AttachToRunningInstance();

        CsiAttachResult AttachToRunningInstance(string instanceId);
    }
}

