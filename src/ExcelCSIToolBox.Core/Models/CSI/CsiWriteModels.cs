using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Models.CSI
{
    public enum CsiProductType
    {
        None = 0,
        ETABS = 1,
        SAP2000 = 2
    }

    public enum CsiMethodRiskLevel
    {
        None = 0,
        Low = 1,
        Medium = 2,
        High = 3,
        Dangerous = 4
    }

    public sealed class CsiWritePreview
    {
        public string OperationName { get; set; }
        public CsiMethodRiskLevel RiskLevel { get; set; }
        public bool RequiresConfirmation { get; set; }
        public bool SupportsDryRun { get; set; }
        public string Summary { get; set; }
        public IReadOnlyList<string> AffectedObjects { get; set; }
    }

    public sealed class CsiMethodDescriptor
    {
        public string ProductType { get; set; }
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string InterfaceName { get; set; }
        public string MethodName { get; set; }
        public IReadOnlyList<CsiParameterDescriptor> Parameters { get; set; }
        public string ReturnType { get; set; }
        public bool IsReadOnly { get; set; }
        public bool IsWrite { get; set; }
        public CsiMethodRiskLevel RiskLevel { get; set; }
        public bool RequiresConfirmation { get; set; }
        public bool SupportsDryRun { get; set; }
        public string ToolName { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
    }

    public sealed class CsiParameterDescriptor
    {
        public string Name { get; set; }
        public string TypeName { get; set; }
        public bool IsOut { get; set; }
        public bool IsOptional { get; set; }
    }
}
