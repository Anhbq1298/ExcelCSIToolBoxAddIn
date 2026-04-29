using System.Collections.Generic;

namespace ExcelCSIToolBox.Data.CSISapModel.Random
{
    public sealed class RandomCsiObjectRequestDto
    {
        public bool AddPoints { get; set; }
        public bool AddFrames { get; set; }
        public bool AddShells { get; set; }
        public int? PointCount { get; set; }
        public int? FrameCount { get; set; }
        public int? ShellCount { get; set; }
        public double? MinX { get; set; }
        public double? MaxX { get; set; }
        public double? MinY { get; set; }
        public double? MaxY { get; set; }
        public double? MinZ { get; set; }
        public double? MaxZ { get; set; }
        public string PointPrefix { get; set; }
        public string FramePrefix { get; set; }
        public string ShellPrefix { get; set; }
        public string FramePropName { get; set; }
        public string ShellPropName { get; set; }
        public int? Seed { get; set; }
    }

    public sealed class RandomCsiObjectResultDto
    {
        public int RequestedPoints { get; set; }
        public int RequestedFrames { get; set; }
        public int RequestedShells { get; set; }
        public int AddedPoints { get; set; }
        public int AddedFrames { get; set; }
        public int AddedShells { get; set; }
        public int FailedItems { get; set; }
        public int Seed { get; set; }
        public List<string> PointNames { get; set; }
        public List<string> FrameNames { get; set; }
        public List<string> ShellNames { get; set; }
        public List<string> FailureReasons { get; set; }
    }
}
