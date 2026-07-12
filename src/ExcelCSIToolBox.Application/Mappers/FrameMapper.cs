using System.Collections.Generic;
using ExcelCSIToolBox.Core.Tabular;

namespace ExcelCSIToolBox.Application.Mappers
{
    public static class CSISapModelFrameDataDataFrameMapper
    {
        public static DataFrame Map(IReadOnlyList<string> frames)
        {
            var rows = new List<IReadOnlyList<object>>();

            if (frames != null)
            {
                foreach (var frame in frames)
                {
                    rows.Add(new object[]
                    {
                        frame
                    });
                }
            }

            var dataFrame = new DataFrame(
                new[] { "UniqueName" },
                rows);
            return dataFrame;
        }
    }
}

