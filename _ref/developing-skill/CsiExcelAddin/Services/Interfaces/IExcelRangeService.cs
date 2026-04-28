using System.Collections.Generic;

namespace CsiExcelAddin.Services.Interfaces
{
    /// <summary>
    /// Reads structured data from a named Excel range or a user-selected range.
    /// Keeping this behind an interface allows unit testing without Excel running.
    /// </summary>
    public interface IExcelRangeReader
    {
        /// <summary>
        /// Reads all non-empty rows from the active selection.
        /// Each row is returned as an ordered list of cell values (as strings).
        /// Returns an empty list — not null — when the selection is empty.
        /// Throws InvalidOperationException when no workbook is open.
        /// </summary>
        IReadOnlyList<IReadOnlyList<string>> ReadSelectedRange();

        /// <summary>
        /// Reads a named range from the active workbook by its defined name.
        /// Returns an empty list when the named range contains no data.
        /// </summary>
        IReadOnlyList<IReadOnlyList<string>> ReadNamedRange(string rangeName);
    }

    /// <summary>
    /// Writes structured data back to a named Excel range or a target cell.
    /// Separating this from reading keeps each service focused and testable.
    /// </summary>
    public interface IExcelRangeWriter
    {
        /// <summary>
        /// Writes rows of values starting at the named range top-left cell.
        /// Existing data in the target area is overwritten.
        /// Throws InvalidOperationException when no workbook is open.
        /// </summary>
        void WriteToNamedRange(string rangeName, IReadOnlyList<IReadOnlyList<object>> data);

        /// <summary>
        /// Writes rows of values starting at a specific cell address (e.g. "B3").
        /// </summary>
        void WriteToCell(string cellAddress, IReadOnlyList<IReadOnlyList<object>> data);

        /// <summary>
        /// Clears all values in the named range without removing formatting.
        /// </summary>
        void ClearNamedRange(string rangeName);
    }
}
