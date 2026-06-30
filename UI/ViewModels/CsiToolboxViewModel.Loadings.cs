using System;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private void GetLoadPatterns()
        {
            var result = _useCases.GetLoadPatterns.Execute();
            if (result.IsSuccess)
            {
                LoadPatterns.Clear();
                if (result.Data != null)
                {
                    foreach (var p in result.Data)
                    {
                        LoadPatterns.Add(p);
                    }
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void DeleteSelectedLoadPatterns(System.Collections.IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0) return;
            
            var list = new System.Collections.Generic.List<string>();
            foreach (var item in selectedItems)
            {
                if (item is ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO dto)
                {
                    list.Add(dto.Name);
                }
            }

            if (list.Count == 0) return;

            var result = _useCases.DeleteLoadPatterns.Execute(list);
            ShowOperationResult(result);
            if (result.IsSuccess)
            {
                GetLoadPatterns(); // refresh list after deletion
            }
        }

        private void GetLoadCombinations()
        {
            var result = _useCases.GetLoadCombinations.Execute();
            if (result.IsSuccess)
            {
                LoadCombinations.Clear();
                if (result.Data != null)
                {
                    foreach (var c in result.Data)
                    {
                        LoadCombinations.Add(c);
                    }
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void DeleteSelectedLoadCombinations(System.Collections.IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0) return;
            
            var list = new System.Collections.Generic.List<string>();
            foreach (var item in selectedItems)
            {
                if (item is ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO dto)
                {
                    list.Add(dto.Name);
                }
            }

            if (list.Count == 0) return;

            var result = _useCases.DeleteLoadCombinations.Execute(list);
            ShowOperationResult(result);
            if (result.IsSuccess)
            {
                GetLoadCombinations(); // refresh list after deletion
            }
        }

        private void ViewLoadCombination(System.Collections.IList selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0) return;
            
            // Only view the first selected item
            var firstItem = selectedItems[0];
            if (firstItem is ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO dto)
            {
                var result = _useCases.GetLoadCombinationDetails.Execute(dto.Name);
                if (result.IsSuccess)
                {
                    var window = new ExcelCSIToolBoxAddIn.UI.Views.LoadCombinationDetailsWindow(result.Data);
                    window.ShowDialog();
                }
                else
                {
                    ShowOperationResult(OperationResult.Failure(result.Message));
                }
            }
        }

        private void ModifyLoadCombinationsInMatrixView()
        {
            var matrixResult = _csiConnectionService.GetLoadCombinationMatrix();
            if (!matrixResult.IsSuccess)
            {
                ShowOperationResult(OperationResult.Failure(matrixResult.Message));
                return;
            }

            var viewModel = new LoadCombinationMatrixViewModel(matrixResult.Data, ProductTitle, _excelOutputService);
            var window = new ExcelCSIToolBoxAddIn.UI.Views.LoadCombinationMatrixView(viewModel);
            window.ShowDialog();

            if (!viewModel.WasSaved)
            {
                return;
            }

            var saveResult = SaveLoadCombinationMatrixChanges(viewModel);
            ShowOperationResult(saveResult);
            if (saveResult.IsSuccess)
            {
                GetLoadCombinations();
            }
        }

        private void ExportLoadCombinationMatrixToExcel()
        {
            var matrixResult = _csiConnectionService.GetLoadCombinationMatrix();
            if (!matrixResult.IsSuccess)
            {
                ShowOperationResult(OperationResult.Failure(matrixResult.Message));
                return;
            }

            var viewModel = new LoadCombinationMatrixViewModel(matrixResult.Data, ProductTitle, _excelOutputService);
            if (viewModel.ExportToExcelRangeCommand.CanExecute(null))
            {
                viewModel.ExportToExcelRangeCommand.Execute(null);
                return;
            }

            ShowOperationResult(OperationResult.Failure("Excel export service is not available."));
        }

        private OperationResult SaveLoadCombinationMatrixChanges(LoadCombinationMatrixViewModel viewModel)
        {
            OperationResult deleteResult = null;
            if (viewModel.SavedDeletedLoadCombinationNames != null && viewModel.SavedDeletedLoadCombinationNames.Count > 0)
            {
                deleteResult = _csiConnectionService.DeleteLoadCombinations(viewModel.SavedDeletedLoadCombinationNames);
                if (!deleteResult.IsSuccess)
                {
                    return deleteResult;
                }
            }

            OperationResult applyResult = null;
            if (viewModel.SavedRows != null && viewModel.SavedRows.Count > 0)
            {
                applyResult = _csiConnectionService.ApplyLoadCombinationMatrix(viewModel.SavedRows);
            }

            if (deleteResult != null && applyResult != null)
            {
                string message = string.Join(" ", new[] { deleteResult.Message, applyResult.Message });
                return applyResult.IsSuccess
                    ? OperationResult.Success(message)
                    : OperationResult.Failure(message);
            }

            if (applyResult != null)
            {
                return applyResult;
            }

            if (deleteResult != null)
            {
                return deleteResult;
            }

            return OperationResult.Success("No load combination changes were saved.");
        }
    }
}
