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

    public sealed class PointGroupAssignmentInfo
    {
        public string PointName { get; set; }
        public IReadOnlyList<string> GroupNames { get; set; }
    }

    public sealed class PointConnectivityInfo
    {
        public string PointName { get; set; }
        public IReadOnlyList<PointConnectedObjectInfo> ConnectedObjects { get; set; }
    }

    public sealed class PointConnectedObjectInfo
    {
        public int ObjectType { get; set; }
        public string ObjectName { get; set; }
        public int PointNumber { get; set; }
    }

    public sealed class PointLocalAxesInfo
    {
        public string PointName { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double C { get; set; }
        public bool Advanced { get; set; }
    }

    public sealed class PointMassInfo
    {
        public string PointName { get; set; }
        public IReadOnlyList<double> MassValues { get; set; }
    }

    public sealed class PointDiaphragmInfo
    {
        public string PointName { get; set; }
        public int DiaphragmOption { get; set; }
        public string DiaphragmName { get; set; }
    }
}
