using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.FrameObject
{
    public sealed class FrameAddRequestDto
    {
        public string UserName { get; set; }
        public string Name { get; set; }
        public string FrameName { get; set; }
        public string UniqueName { get; set; }
        public string PropName { get; set; }
        public string SectionName { get; set; }
        public string PointIName { get; set; }
        public string PointJName { get; set; }
        public double? Xi { get; set; }
        public double? Yi { get; set; }
        public double? Zi { get; set; }
        public double? Xj { get; set; }
        public double? Yj { get; set; }
        public double? Zj { get; set; }
    }

    public sealed class FrameAddBatchRequestDto
    {
        public List<FrameAddRequestDto> Frames { get; set; }
    }

    public sealed class FrameAddResultDto
    {
        public bool Success { get; set; }
        public string FrameName { get; set; }
        public string AddMethod { get; set; }
        public string FailureReason { get; set; }
        public int? ReturnCode { get; set; }
    }

    public sealed class FrameAddBatchResultDto
    {
        public int TotalRequested { get; set; }
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public List<string> SuccessfulFrameNames { get; set; }
        public List<FrameAddResultDto> FailedItems { get; set; }
        public List<FrameAddResultDto> Results { get; set; }
    }

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
