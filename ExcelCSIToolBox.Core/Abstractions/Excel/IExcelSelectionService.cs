using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Core.Abstractions.Excel
{
    public interface IExcelSelectionService
    {
        OperationResult<IReadOnlyList<string>> ReadSingleColumnTextValues();

        OperationResult<IReadOnlyList<ExcelPointCartesianRow>> ReadPointCartesianRows();

        OperationResult<IReadOnlyList<ExcelFrameByCoordRow>> ReadFrameByCoordRows();

        OperationResult<IReadOnlyList<ExcelFrameByPointRow>> ReadFrameByPointRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelSteelISectionRow>> ReadSteelISectionRows();
        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelSteelChannelSectionRow>> ReadSteelChannelSectionRows();
        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelSteelAngleSectionRow>> ReadSteelAngleSectionRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelSteelPipeSectionRow>> ReadSteelPipeSectionRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelSteelTubeSectionRow>> ReadSteelTubeSectionRows();

        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelConcreteRectangleSectionRow>> ReadConcreteRectangleSectionRows();
        OperationResult<IReadOnlyList<ExcelCSIToolBox.Core.Tabular.ExcelConcreteCircleSectionRow>> ReadConcreteCircleSectionRows();
    }
}


