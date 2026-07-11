using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Common.Commands;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private void RefreshCommandStates()
        {
            var commands = new ICommand[]
            {
                RefreshRunningCsiInstancesCommand,
                AttachToRunningCsiCommand,
                CloseCurrentInstanceCommand,
                ToggleModelLockCommand,
                SelectWorkspacePageCommand,
                RefreshFrameStiffnessSectionsCommand,
                RefreshAreaStiffnessSectionsCommand,
                SelectVisibleFrameStiffnessSectionsCommand,
                ClearFrameStiffnessSectionSelectionCommand,
                SelectVisibleAreaStiffnessSectionsCommand,
                ClearAreaStiffnessSectionSelectionCommand,
                ApplyFrameStiffnessModifiersCommand,
                ApplyAreaStiffnessModifiersCommand,
                ResetFrameStiffnessModifiersCommand,
                ResetAreaStiffnessModifiersCommand,
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
                ExportLoadCombinationMatrixToExcelCommand,
                AddLoadCombinationFromExcelCommand,
                DeleteSelectedLoadCombinationsCommand,
                ViewLoadCombinationCommand,
                ExportAnalysisResultTableCommand,
                ExportEtabsTableItemCommand,
                GetBaseReactionsCommand,
                GetModalMassParticipationRatiosCommand,
                GetStoryForcesCommand,
                GetStoryDriftsCommand,
                GetStoryMaxOverAverageDisplacementsCommand,
                GetStoryMaxOverAverageDriftsCommand,
                GetMassSummaryByStoryCommand,
                GetFrameSectionsCommand,
                EditFrameSectionCommand,
                OpenCreateSectionDialogCommand,
                OpenShellUniformLoadSetFormCommand,
                ExportShellUniformLoadSetDefinitionsCommand,
                OpenCreateArrayPerpendicularToPathWindowCommand,
                OpenArrayBetweenTwoLinesWindowCommand,
                PickPoint1Command,
                PickPoint2Command,
                PickReferenceFrameCommand,
                PickLine1Command,
                PickLine2Command,
                CreateFramesCommand,
                CreateArrayBetweenTwoLinesFramesCommand,
                OpenOffsetFromSetOfLinesCommand,
                OffsetGetSelectedLinesCommand,
                OffsetPreviewCommand,
                OffsetCreateInEtabsCommand,
                OffsetClearCommand,
                OffsetRefreshSectionsCommand,
                CloseWindowCommand,
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
            if (ProductType == CsiProductType.SAP2000)
            {
                IsConnected = false;
                HasRunningCsiInstance = false;
                StatusText = "Not attached";
                return;
            }

            RefreshRunningCsiInstances();

            OperationResult<CSISapModelConnectionInfoDTO> result = SelectedRunningCsiInstance == null
                ? _useCases.LoadConnection.Execute()
                : _csiConnectionService.AttachToRunningInstance(SelectedRunningCsiInstance.InstanceId);

            ApplyConnectionResult(result, showMessage);
        }

        private void RefreshRunningCsiInstancesFromUi()
        {
            RefreshRunningCsiInstances();
            if (RunningCsiInstances.Count == 0)
            {
                StatusText = $"No running {_productName} instance found.";
            }
            else
            {
                StatusText = $"Found {RunningCsiInstances.Count} running {_productName} instance(s).";
                if (SelectedRunningCsiInstance != null)
                {
                    AttachToSelectedRunningInstance(SelectedRunningCsiInstance, showMessage: false);
                }
                else
                {
                    SelectedRunningCsiInstance = RunningCsiInstances[0];
                    AttachToSelectedRunningInstance(SelectedRunningCsiInstance, showMessage: false);
                }
            }
        }

        private void AttachToSelectedRunningInstance(CsiRunningInstanceViewModel instance, bool showMessage)
        {
            if (instance == null)
            {
                return;
            }

            OperationResult<CSISapModelConnectionInfoDTO> result = _csiConnectionService.AttachToRunningInstance(instance.InstanceId);
            ApplyConnectionResult(result, showMessage);
        }

        private void ApplyConnectionResult(OperationResult<CSISapModelConnectionInfoDTO> result, bool showMessage)
        {
            if (result.IsSuccess && result.Data != null)
            {
                IsConnected = true;
                HasRunningCsiInstance = true;
                ModelName = string.IsNullOrWhiteSpace(result.Data.ModelFileName)
                    ? "Unknown model"
                    : result.Data.ModelFileName;
                ModelPath = result.Data.ModelPath ?? string.Empty;
                CurrentModelUnitText = string.IsNullOrWhiteSpace(result.Data.ModelCurrentUnit)
                    ? "Units unavailable"
                    : result.Data.ModelCurrentUnit;
                SyncSelectedUnitSystemFromEtabs();
                RefreshModelLockState();
                StatusText = "Attached successfully.";
                SelectRunningInstanceFromConnection(result.Data);
                
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
            IsModelLocked = false;
            SetDetachedModelInfo("Not yet attached");
            StatusText = string.IsNullOrWhiteSpace(result.Message)
                ? $"{_productName} connection unavailable."
                : result.Message;
                
            LoadPatterns.Clear();
            LoadCombinations.Clear();
            FrameSections.Clear();
            FrameStiffnessSections.Clear();
            AreaStiffnessSections.Clear();
            OffsetAvailableSections.Clear();
            OffsetSelectedSection = null;
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

        private void RefreshRunningCsiInstances()
        {
            if (ProductType == CsiProductType.SAP2000)
            {
                RunningCsiInstances.Clear();
                HasRunningCsiInstance = false;
                return;
            }

            string selectedInstanceId = SelectedRunningCsiInstance == null ? null : SelectedRunningCsiInstance.InstanceId;
            OperationResult<IReadOnlyList<CSISapModelRunningInstanceDTO>> result = _csiConnectionService.GetRunningInstances();
            if (!result.IsSuccess)
            {
                HasRunningCsiInstance = false;
                return;
            }

            try
            {
                _isRefreshingRunningCsiInstances = true;
                RunningCsiInstances.Clear();
                if (result.Data != null)
                {
                    foreach (CSISapModelRunningInstanceDTO instance in result.Data)
                    {
                        RunningCsiInstances.Add(new CsiRunningInstanceViewModel
                        {
                            InstanceId = instance.InstanceId,
                            ProcessId = instance.ProcessId,
                            DisplayName = string.IsNullOrWhiteSpace(instance.DisplayName)
                                ? instance.ModelFileName
                                : instance.DisplayName,
                            ModelPath = instance.ModelPath,
                            ModelFileName = instance.ModelFileName,
                            ModelCurrentUnit = instance.ModelCurrentUnit
                        });
                    }
                }

                SelectedRunningCsiInstance = FindRunningInstanceById(selectedInstanceId)
                    ?? (RunningCsiInstances.Count == 1 ? RunningCsiInstances[0] : null);

                HasRunningCsiInstance = RunningCsiInstances.Count > 0;
            }
            finally
            {
                _isRefreshingRunningCsiInstances = false;
            }
        }

        private void SelectRunningInstanceFromConnection(CSISapModelConnectionInfoDTO connectionInfo)
        {
            if (connectionInfo == null)
            {
                return;
            }

            CsiRunningInstanceViewModel matchingInstance = null;
            if (connectionInfo.ProcessId.HasValue)
            {
                foreach (CsiRunningInstanceViewModel instance in RunningCsiInstances)
                {
                    if (instance.ProcessId == connectionInfo.ProcessId)
                    {
                        matchingInstance = instance;
                        break;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(connectionInfo.ModelPath))
            {
                foreach (CsiRunningInstanceViewModel instance in RunningCsiInstances)
                {
                    if (matchingInstance == null &&
                        string.Equals(instance.ModelPath, connectionInfo.ModelPath, StringComparison.OrdinalIgnoreCase))
                    {
                        matchingInstance = instance;
                        break;
                    }
                }
            }

            if (matchingInstance == null && RunningCsiInstances.Count == 1)
            {
                matchingInstance = RunningCsiInstances[0];
            }

            if (matchingInstance == null)
            {
                return;
            }

            try
            {
                _isRefreshingRunningCsiInstances = true;
                SelectedRunningCsiInstance = matchingInstance;
            }
            finally
            {
                _isRefreshingRunningCsiInstances = false;
            }
        }

        private CsiRunningInstanceViewModel FindRunningInstanceById(string instanceId)
        {
            if (string.IsNullOrWhiteSpace(instanceId))
            {
                return null;
            }

            foreach (CsiRunningInstanceViewModel instance in RunningCsiInstances)
            {
                if (string.Equals(instance.InstanceId, instanceId, StringComparison.OrdinalIgnoreCase))
                {
                    return instance;
                }
            }

            return null;
        }

        private void CloseCurrentInstance()
        {
            var result = _useCases.CloseCurrentInstance.Execute();

            if (result.IsSuccess)
            {
                IsConnected = false;
                IsModelLocked = false;
                SetDetachedModelInfo("Not connected");
                StatusText = result.Message;

                LoadPatterns.Clear();
                LoadCombinations.Clear();
                FrameSections.Clear();
                FrameStiffnessSections.Clear();
                AreaStiffnessSections.Clear();
                OffsetAvailableSections.Clear();
                OffsetSelectedSection = null;
                SelectedFrameSection = null;
                RefreshRunningCsiInstances();

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
            CurrentModelUnitText = SelectedUnitSystem == null ? "Not yet attached" : SelectedUnitSystem.PresentUnitsText;
        }

        private void RefreshModelLockState()
        {
            OperationResult<bool> lockResult = _csiConnectionService.GetModelIsLocked();
            if (lockResult.IsSuccess)
            {
                IsModelLocked = lockResult.Data;
            }
        }

        private void ToggleModelLock()
        {
            bool nextLockState = !IsModelLocked;
            OperationResult result = _csiConnectionService.SetModelIsLocked(nextLockState);
            if (result.IsSuccess)
            {
                IsModelLocked = nextLockState;
                StatusText = result.Message;
                return;
            }

            MessageBox.Show(
                result.Message,
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            StatusText = result.Message;
            RefreshModelLockState();
        }

        private System.Collections.ObjectModel.ObservableCollection<EtabsUnitSystem> CreateAvailableUnitSystems()
        {
            return new System.Collections.ObjectModel.ObservableCollection<EtabsUnitSystem>
            {
                new EtabsUnitSystem("kN-m", 4, 6, 2, "kN", "kN-m", "m", 6),
                new EtabsUnitSystem("N-mm", 3, 4, 2, "N", "N-mm", "mm", 9),
                new EtabsUnitSystem("kip-ft", 2, 2, 2, "kip", "kip-ft", "ft", 4),
                new EtabsUnitSystem("lb-in", 1, 1, 2, "lb", "lb-in", "in", 1)
            };
        }

        private void SyncSelectedUnitSystemFromEtabs()
        {
            OperationResult<CSISapModelPresentUnitSystemDTO> result = _csiConnectionService.GetPresentUnitSystem();
            EtabsUnitSystem matchedUnitSystem = null;
            if (result.IsSuccess)
            {
                foreach (EtabsUnitSystem unitSystem in AvailableUnitSystems)
                {
                    if (unitSystem.Matches(result.Data))
                    {
                        matchedUnitSystem = unitSystem;
                        break;
                    }
                }
            }

            if (matchedUnitSystem == null && AvailableUnitSystems.Count > 0)
            {
                matchedUnitSystem = AvailableUnitSystems[0];
            }

            _isInitializingUnitSystems = true;
            SelectedUnitSystem = matchedUnitSystem;
            _isInitializingUnitSystems = false;
            CurrentModelUnitText = matchedUnitSystem == null ? "Units unavailable" : matchedUnitSystem.PresentUnitsText;
        }

        private bool PrepareExportWithGlobalUnit()
        {
            return ApplySelectedGlobalUnit(showMessages: true);
        }

        private bool ApplySelectedGlobalUnit(bool showMessages)
        {
            if (_isApplyingGlobalUnit)
            {
                return true;
            }

            if (SelectedUnitSystem == null)
            {
                if (showMessages)
                {
                    MessageBox.Show("Please select a unit system first.", ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                return false;
            }

            try
            {
                _isApplyingGlobalUnit = true;
                _etabsUnitService.SetPresentUnitsFromMainWindow();

                CurrentModelUnitText = SelectedUnitSystem.PresentUnitsText;
                StatusText = "Unit system set to " + SelectedUnitSystem.DisplayName + ".";
                return true;
            }
            catch (InvalidOperationException ex)
            {
                if (showMessages)
                {
                    MessageBox.Show(ex.Message, ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                StatusText = ex.Message;
                return false;
            }
            finally
            {
                _isApplyingGlobalUnit = false;
            }
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
