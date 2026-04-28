namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters
{
    public interface ICsiModelAdapter
    {
        string ApplicationName { get; }

        CsiAttachResult AttachToRunningInstance();
    }
}

