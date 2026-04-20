using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public interface IExcelSelectionService
    {
        OperationResult<IReadOnlyList<string>> ReadSingleColumnTextValues();

        OperationResult<IReadOnlyList<ExcelPointCartesianRow>> ReadPointCartesianRows();

        OperationResult<IReadOnlyList<ExcelFrameByCoordRow>> ReadFrameByCoordRows();

        OperationResult<IReadOnlyList<ExcelFrameByPointRow>> ReadFrameByPointRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelISectionRow>> ReadSteelISectionRows();
        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelChannelSectionRow>> ReadSteelChannelSectionRows();
        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelAngleSectionRow>> ReadSteelAngleSectionRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelPipeSectionRow>> ReadSteelPipeSectionRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelTubeSectionRow>> ReadSteelTubeSectionRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelConcreteRectangleSectionRow>> ReadConcreteRectangleSectionRows();
        OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelConcreteCircleSectionRow>> ReadConcreteCircleSectionRows();
    }
}
