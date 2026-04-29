using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.PointObject
{
    public sealed class PointObjectInfo
    {
        public string Name { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public bool IsSelected { get; set; }
        public string CoordinateSystem { get; set; }
    }

    public sealed class PointLoadInfo
    {
        public string PointName { get; set; }
        public string LoadPattern { get; set; }
        public double F1 { get; set; }
        public double F2 { get; set; }
        public double F3 { get; set; }
        public double M1 { get; set; }
        public double M2 { get; set; }
        public double M3 { get; set; }
        public string CoordinateSystem { get; set; }
    }

    public sealed class PointRestraintInfo
    {
        public string PointName { get; set; }
        public bool U1 { get; set; }
        public bool U2 { get; set; }
        public bool U3 { get; set; }
        public bool R1 { get; set; }
        public bool R2 { get; set; }
        public bool R3 { get; set; }
    }

    public sealed class PointSpringInfo
    {
        public string PointName { get; set; }
        public IReadOnlyList<double> Stiffness { get; set; }
        public string CoordinateSystem { get; set; }
    }
}
