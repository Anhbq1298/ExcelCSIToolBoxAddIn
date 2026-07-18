using System.Collections.Generic;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IDropPanelEtabsService
    {
        OperationResult<DropPanelModelContext> GetModelContext();

        OperationResult<IReadOnlyList<string>> GetConcreteMaterialNames();

        /// <summary>
        /// Resolves the ETABS user-friendly label for each supplied frame unique name.
        /// Returns a dictionary mapping unique name → label. Unknown names map to an
        /// empty string. Used by callers that need to pre-filter by label prefix before
        /// invoking ReadColumns.
        /// </summary>
        OperationResult<IReadOnlyDictionary<string, string>> GetFrameLabels(
            IReadOnlyList<string> frameUniqueNames);

        OperationResult<IReadOnlyList<DropPanelColumnInfo>> ReadColumns(
            IReadOnlyList<string> frameNames,
            double verticalRatioTolerance);

        OperationResult<DropPanelPreparationSnapshot> PrepareSnapshot(
            IReadOnlyList<DropPanelColumnInfo> columns,
            IReadOnlyList<DropPanelRequest> requests,
            DropPanelOptions options);

        OperationResult<DropPanelApplyResult> Apply(DropPanelOperationPlan plan, DropPanelOptions options);
    }
}
