using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Core.Tabular;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public static class CsiFrameDataDataFrameMapper
    {
        public static DataFrame Map(IReadOnlyList<string> frameUniqueNames)
        {
            var rows = new List<IReadOnlyList<object>>();

            if (frameUniqueNames != null)
            {
                foreach (var frameUniqueName in frameUniqueNames)
                {
                    rows.Add(new object[] { frameUniqueName });
                }
            }

            return new DataFrame(
                new[] { "FrameUniqueName" },
                rows);
        }
    }
}
