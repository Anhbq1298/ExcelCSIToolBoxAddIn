using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;
using System.Collections.Generic;

namespace ExcelCSIToolBox.Data
{
    #region Results & Data DTOs
    
    public class CSISapModelAddFramesResultDTO
    {
        public int AddedCount { get; set; }
        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }

    public class CSISapModelAddPointsResultDTO
    {
        public int AddedCount { get; set; }
        public IReadOnlyList<string> FailedRowMessages { get; set; }
    }

    /// <summary>
    /// Value object containing CSI connection details needed by the UI.
    /// </summary>
    public class CSISapModelConnectionInfoDTO
    {
        public bool IsConnected { get; set; }
        public string ModelPath { get; set; }
        public string ModelFileName { get; set; }
        public string ModelCurrentUnit { get; set; }
        /// <summary>
        /// Optional COM object references for future CSI operations.
        /// </summary>
        public object CsiObject { get; set; }
        public object SapModel { get; set; }
    }

    public class CSISapModelPointDataDTO
    {
        public string PointUniqueName { get; set; }
        public string PointLabel { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    public class CSISapModelStatisticsDTO
    {
        public int PointCount { get; set; }
        public int FrameCount { get; set; }
        public int ShellCount { get; set; }
        public int LoadPatternCount { get; set; }
        public int LoadCombinationCount { get; set; }
    }

    #endregion

    #region Input / Table Format Classes

    public class CSISapModelPointCartesianInput
    {
        public int ExcelRowNumber { get; set; }
        public string UniqueName { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    public class CSISapModelFrameByCoordInput
    {
        public int ExcelRowNumber { get; set; }
        public string UniqueName { get; set; }
        public string SectionName { get; set; }
        public double Xi { get; set; }
        public double Yi { get; set; }
        public double Zi { get; set; }
        public double Xj { get; set; }
        public double Yj { get; set; }
        public double Zj { get; set; }
    }

    public class CSISapModelFrameByPointInput
    {
        public int ExcelRowNumber { get; set; }
        public string UniqueName { get; set; }
        public string SectionName { get; set; }
        public string Point1Name { get; set; }
        public string Point2Name { get; set; }
    }

    public class CSISapModelSteelISectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CSISapModelSteelChannelSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CSISapModelSteelAngleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double Tw { get; set; }
        public double Tf { get; set; }
    }

    public class CSISapModelSteelPipeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double OutsideDiameter { get; set; }
        public double WallThickness { get; set; }
    }

    public class CSISapModelSteelTubeSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
        public double T { get; set; }
    }

    public class CSISapModelConcreteRectangleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double H { get; set; }
        public double B { get; set; }
    }

    public class CSISapModelConcreteCircleSectionInput
    {
        public string SectionName { get; set; }
        public string MaterialName { get; set; }
        public double D { get; set; }
    }

    #endregion

    #region Constants

    public static class CSISapModelObjectTypeIds
    {
        public const int Point = 1;
        public const int Frame = 2;
        public const int Shell = 5;
    }

    #endregion
}


