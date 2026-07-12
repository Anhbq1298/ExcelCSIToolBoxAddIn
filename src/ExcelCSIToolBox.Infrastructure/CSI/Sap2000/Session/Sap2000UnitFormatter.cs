namespace ExcelCSIToolBox.Infrastructure.CSI.Sap2000.Session
{
    public static class Sap2000UnitFormatter
    {
        public static string FormatPresentUnits(SAP2000v1.eUnits units)
        {
            return units.ToString().Replace("_", "-");
        }
    }
}

