using System.Collections.Generic;
using ExcelCSIToolBox.Core.Tabular;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.Mappers
{
    public static class CSISapModelPointDataDataFrameMapper
    {
        public static DataFrame Map(IReadOnlyList<CSISapModelPointDataDTO> points)
        {
            var rows = new List<IReadOnlyList<object>>();

            if (points != null)
            {
                foreach (var point in points)
                {
                    rows.Add(new object[]
                    {
                        point.PointUniqueName,
                        point.PointLabel,
                        point.X,
                        point.Y,
                        point.Z
                    });
                }
            }

            var dataFrame = new DataFrame(
                new[] { "UniqueName", "Point Label", "X", "Y", "Z" },
                rows);
            return dataFrame;
        }
    }
}

