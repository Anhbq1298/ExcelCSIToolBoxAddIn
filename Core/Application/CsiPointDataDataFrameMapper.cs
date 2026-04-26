using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Core.Tabular;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public static class CsiPointDataDataFrameMapper
    {
        public static DataFrame Map(IReadOnlyList<CsiPointData> points)
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
