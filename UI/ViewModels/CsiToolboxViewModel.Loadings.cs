using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Infrastructure.Excel;

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

        private void OpenShellUniformLoadSetForm()
        {
            using (var form = new ExcelCSIToolBoxAddIn.UI.Forms.ShellUniformLoadSetForm(
                _csiConnectionService,
                ownerHandle => ExportShellUniformLoadSetDefinitions(ownerHandle)))
            {
                System.Windows.Forms.IWin32Window owner = null;
                try
                {
                    var excelApp = ExcelApplicationProvider.GetApplication();
                    if (excelApp != null)
                    {
                        owner = new ExcelCSIToolBoxAddIn.UI.Forms.Win32WindowWrapper(new IntPtr(excelApp.Hwnd));
                    }
                }
                catch
                {
                    // Ignore wrapper creation errors
                }

                if (owner != null)
                {
                    form.ShowDialog(owner);
                }
                else
                {
                    form.ShowDialog();
                }
            }
        }

        private void ExportShellUniformLoadSetDefinitions()
        {
            ExportShellUniformLoadSetDefinitions(IntPtr.Zero);
        }

        private void ExportShellUniformLoadSetDefinitions(IntPtr ownerHandle)
        {
            var result = _csiConnectionService.GetShellUniformLoadSetDefinitions();
            if (!result.IsSuccess)
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
                return;
            }

            IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions = result.Data ?? new List<ShellUniformLoadSetDefinitionDto>();
            if (definitions.Count == 0)
            {
                ShowOperationResult(OperationResult.Failure("No Shell Uniform Load Set definitions were found in the active ETABS model."));
                return;
            }

            OutputTableExportWorkflow.Run(
                new OutputTableExportConfig
                {
                    TableDisplayName = "Shell Uniform Load Set Definitions",
                    Breadcrumb = "ETABS Toolbox / Model / Shell Uniform Load Set Manager / Export Current Definitions",
                    Description = "Select the Excel anchor cell and output format for current Shell Uniform Load Set definitions.",
                    PopupProfileKey = "EtabsObjectConnectivity",
                    EmptyDataMessage = "No Shell Uniform Load Set definitions were found in the active ETABS model.",
                    WorksheetNamePrefix = "Shell Uniform Load Set",
                    DefaultAddHeaders = true,
                    StaticRecordCount = definitions.Count,
                    StaticSuccessMessage = "Exported " + definitions.Count.ToString(CultureInfo.InvariantCulture) + " Shell Uniform Load Set definition(s) to Excel.",
                    StaticExportValuesFactory = addHeaders => CreateShellUniformLoadSetExportValues(definitions, addHeaders)
                },
                _useCases,
                _csiConnectionService,
                _excelOutputService,
                ownerHandle == IntPtr.Zero ? GetActiveOwnerWindow() : null,
                ownerHandle);
        }

        private static object[,] CreateShellUniformLoadSetExportValues(IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions, bool addHeaders)
        {
            var patternNames = definitions
                .Where(definition => definition != null && definition.LoadValuesByPattern != null)
                .SelectMany(definition => definition.LoadValuesByPattern.Keys)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int dataRowOffset = addHeaders ? 1 : 0;
            int rowCount = definitions.Count + dataRowOffset;
            int columnCount = patternNames.Count + 1;
            var values = new object[rowCount, columnCount];
            if (addHeaders)
            {
                values[0, 0] = "UniformLoadSetName";
                for (int columnIndex = 0; columnIndex < patternNames.Count; columnIndex++)
                {
                    values[0, columnIndex + 1] = patternNames[columnIndex];
                }
            }

            for (int rowIndex = 0; rowIndex < definitions.Count; rowIndex++)
            {
                ShellUniformLoadSetDefinitionDto definition = definitions[rowIndex];
                int outputRowIndex = rowIndex + dataRowOffset;
                values[outputRowIndex, 0] = definition == null ? string.Empty : definition.Name ?? string.Empty;
                for (int columnIndex = 0; columnIndex < patternNames.Count; columnIndex++)
                {
                    double loadValue;
                    if (definition != null &&
                        definition.LoadValuesByPattern != null &&
                        definition.LoadValuesByPattern.TryGetValue(patternNames[columnIndex], out loadValue))
                    {
                        values[outputRowIndex, columnIndex + 1] = loadValue;
                    }
                }
            }

            return values;
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
