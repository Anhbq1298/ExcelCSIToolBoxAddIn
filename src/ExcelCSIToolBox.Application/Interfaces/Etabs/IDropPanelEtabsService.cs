using System.Collections.Generic;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IDropPanelEtabsService
    {
        OperationResult<DropPanelModelContext> GetModelContext();

        OperationResult<IReadOnlyList<string>> GetDropPropertyNames();

        OperationResult<IReadOnlyList<DropPanelColumnInfo>> ReadSelectedColumns(double verticalRatioTolerance);

        OperationResult<DropPanelPreparationSnapshot> PrepareSnapshot(
            IReadOnlyList<DropPanelColumnInfo> columns,
            IReadOnlyList<DropPanelRequest> requests,
            DropPanelOptions options);

        OperationResult HighlightAreas(IReadOnlyList<string> areaNames);

        OperationResult<DropPanelApplyResult> Apply(DropPanelOperationPlan plan, DropPanelOptions options);

        OperationResult Rollback();

        bool IsRollbackAvailable { get; }
    }
}
