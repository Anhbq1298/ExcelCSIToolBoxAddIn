namespace ExcelCSIToolBoxAddIn.Adapters
{
    public interface ICsiModelAdapter
    {
        string ApplicationName { get; }

        CsiAttachResult AttachToRunningInstance();
    }
}
