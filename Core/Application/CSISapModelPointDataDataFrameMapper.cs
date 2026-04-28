using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Core.Tabular;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public static class CSISapModelPointDataDataFrameMapper
    {
        public static DataFrame Map(IReadOnlyList<CSISapModelPointData> points)
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

            return new DataFrame(
                new[] { "UniqueName", "Point Label", "X", "Y", "Z" },
                rows);
        }
    }
}
