using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelCSIToolBoxAddIn.Core.Tabular
{
    public class DataFrame
    {
        public DataFrame(IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<object>> rows)
        {
            if (columns == null || columns.Count == 0)
            {
                throw new ArgumentException("DataFrame must contain at least one column.", nameof(columns));
            }

            Columns = new ReadOnlyCollection<string>(columns.ToList());
            Rows = new ReadOnlyCollection<IReadOnlyList<object>>((rows ?? Array.Empty<IReadOnlyList<object>>()).ToList());
        }

        public IReadOnlyList<string> Columns { get; }

        public IReadOnlyList<IReadOnlyList<object>> Rows { get; }
    }
}
