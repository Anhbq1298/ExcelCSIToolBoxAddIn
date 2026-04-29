using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.FrameObject
{
    public sealed class FrameObjectInfo
    {
        public string Name { get; set; }
        public string PointI { get; set; }
        public string PointJ { get; set; }
        public string SectionName { get; set; }
        public bool IsSelected { get; set; }
    }

    public sealed class FrameEndPointInfo
    {
        public string FrameName { get; set; }
        public string PointI { get; set; }
        public string PointJ { get; set; }
    }

    public sealed class FrameSectionInfo
    {
        public string FrameName { get; set; }
        public string SectionName { get; set; }
        public string AutoSelectList { get; set; }
    }

    public sealed class FrameReleaseInfo
    {
        public string FrameName { get; set; }
        public IReadOnlyList<bool> StartReleases { get; set; }
        public IReadOnlyList<bool> EndReleases { get; set; }
        public IReadOnlyList<double> StartSpringValues { get; set; }
        public IReadOnlyList<double> EndSpringValues { get; set; }
    }

    public sealed class FrameLoadInfo
    {
        public string FrameName { get; set; }
        public string LoadPattern { get; set; }
        public string LoadType { get; set; }
        public int Direction { get; set; }
        public double Value1 { get; set; }
        public double Value2 { get; set; }
        public double Distance1 { get; set; }
        public double Distance2 { get; set; }
        public bool IsRelativeDistance { get; set; }
        public string CoordinateSystem { get; set; }
    }
}
