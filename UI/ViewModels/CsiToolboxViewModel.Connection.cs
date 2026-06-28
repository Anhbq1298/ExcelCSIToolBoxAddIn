using System;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Common.Commands;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private void RefreshCommandStates()
        {
            var commands = new ICommand[]
            {
                CloseCurrentInstanceCommand,
                CreateIshapeSectionCommand,
                CreateChannelSectionCommand,
                CreateAngleSectionCommand,
                CreateTubeSectionCommand,
                CreatePipeSectionCommand,
                CreateConcreteRectangleSectionCommand,
                CreateConcreteCircleSectionCommand,
                SelectPointsByUniqueNameCommand,
                SelectFramesByUniqueNameCommand,
                AddPointByCartesianCommand,
                SetPointsCommand,
                RenameSelectedPointsCommand,
                GetSelectedPointsCommand,
                AddFramesByCoordinatesCommand,
                AddFramesByPointNamesCommand,
                SetFramesCommand,
                RenameFramesCommand,
                GetSelectedFramesCommand,
                GetFrameSectionPropertyCommand,
                SetFrameSectionPropertyCommand,
                GetFrameGroupAssignmentCommand,
                SetFrameGroupAssignmentCommand,
                GetFrameModifierCommand,
                SetFrameModifierCommand,
                CreateShellAreasFromSelectedFramesCommand,
                GetPointGroupAssignmentCommand,
                SetPointGroupAssignmentCommand,
                GetLoadPatternsCommand,
                AddLoadPatternFromExcelCommand,
                DeleteSelectedLoadPatternsCommand,
                GetLoadCombinationsCommand,
                ModifyLoadCombinationsInMatrixViewCommand,
                AddLoadCombinationFromExcelCommand,
                DeleteSelectedLoadCombinationsCommand,
                ViewLoadCombinationCommand,
                ExportAnalysisResultTableCommand,
                GetBaseReactionsCommand,
                GetModalMassParticipationRatiosCommand,
                GetStoryForcesCommand,
                GetStoryDriftsCommand,
                GetStoryMaxOverAverageDisplacementsCommand,
                GetStoryMaxOverAverageDriftsCommand,
                GetMassSummaryByStoryCommand,
                GetFrameSectionsCommand,
                EditFrameSectionCommand
            };

            foreach (ICommand command in commands)
            {
                if (command is IRelayCommand relayCommand)
                {
                    relayCommand.RaiseCanExecuteChanged();
                }
            }

            CommandManager.InvalidateRequerySuggested();
        }

        private void LoadConnectionState(bool showMessage)
        {
            var result = _useCases.LoadConnection.Execute();

            if (result.IsSuccess && result.Data != null)
            {
                IsConnected = true;
                ModelName = string.IsNullOrWhiteSpace(result.Data.ModelFileName)
                    ? "Unknown model"
                    : result.Data.ModelFileName;
                ModelPath = result.Data.ModelPath ?? string.Empty;
                CurrentModelUnitText = string.IsNullOrWhiteSpace(result.Data.ModelCurrentUnit)
                    ? "Units unavailable"
                    : result.Data.ModelCurrentUnit;
                StatusText = "Attached successfully.";
                
                // Automatically refresh lists when connection is established
                GetLoadPatterns();
                GetLoadCombinations();
                GetFrameSections();
                RefreshCommandStates();
                if (showMessage)
                {
                    ShowOperationResult(OperationResult.Success("Successfully attached to the running application."));
                }

                return;
            }

            IsConnected = false;
            SetDetachedModelInfo("Not yet attached");
            StatusText = string.IsNullOrWhiteSpace(result.Message)
                ? $"{_productName} connection unavailable."
                : result.Message;
                
            LoadPatterns.Clear();
            LoadCombinations.Clear();
            FrameSections.Clear();
            SelectedFrameSection = null;

            if (showMessage)
            {
                MessageBox.Show(
                    StatusText,
                    ProductTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void CloseCurrentInstance()
        {
            var result = _useCases.CloseCurrentInstance.Execute();

            if (result.IsSuccess)
            {
                IsConnected = false;
                SetDetachedModelInfo("Not connected");
                StatusText = result.Message;

                LoadPatterns.Clear();
                LoadCombinations.Clear();
                FrameSections.Clear();
                SelectedFrameSection = null;

                MessageBox.Show(
                    result.Message,
                    ProductTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            MessageBox.Show(
                result.Message,
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            StatusText = result.Message;
        }

        private void SetDetachedModelInfo(string modelNameText)
        {
            ModelName = modelNameText;
            ModelPath = string.Empty;
            CurrentModelUnitText = "Not yet attached";
        }

        private void ShowOperationResult(OperationResult result)
        {
            MessageBox.Show(
                result.Message,
                ProductTitle,
                MessageBoxButton.OK,
                result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private void ShowPlaceholder(string featureName)
        {
            MessageBox.Show(
                $"{featureName} is a placeholder for phase 1.",
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
