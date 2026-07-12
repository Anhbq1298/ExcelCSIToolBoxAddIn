using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ExcelCSIToolBoxAddIn.UI.Common.Commands;
using ExcelCSIToolBox.Core.Models.AnalysisResults;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.AnalysisResults;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Connectivity;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.DatabaseTables.MiscellaneousData;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public sealed class AnalysisResultsViewModel : ViewModelBase
    {
        private readonly Func<bool> _canUseActiveModel;
        private readonly Func<bool> _canExecuteEtabsAction;
        private string _activeTableCategory;
        private string _activeAnalysisResultsGroup;
        private string _selectedAnalysisResultTable;

        public AnalysisResultsViewModel(
            Func<bool> canUseActiveModel,
            Func<bool> canExecuteEtabsAction,
            Action<AnalysisResultItem> runAnalysisResult,
            Action<object> runEtabsTableItem,
            Action openGetBaseReactionsDialog,
            Action openModalMassParticipationRatiosDialog,
            Action openStoryForcesDialog,
            Action openStoryDriftsDialog,
            Action openStoryMaxOverAverageDisplacementsDialog,
            Action openStoryMaxOverAverageDriftsDialog,
            Action openMassSummaryByStoryDialog)
        {
            _canUseActiveModel = canUseActiveModel ?? throw new ArgumentNullException(nameof(canUseActiveModel));
            _canExecuteEtabsAction = canExecuteEtabsAction ?? throw new ArgumentNullException(nameof(canExecuteEtabsAction));
            if (runAnalysisResult == null) throw new ArgumentNullException(nameof(runAnalysisResult));
            if (runEtabsTableItem == null) throw new ArgumentNullException(nameof(runEtabsTableItem));
            if (openGetBaseReactionsDialog == null) throw new ArgumentNullException(nameof(openGetBaseReactionsDialog));
            if (openModalMassParticipationRatiosDialog == null) throw new ArgumentNullException(nameof(openModalMassParticipationRatiosDialog));
            if (openStoryForcesDialog == null) throw new ArgumentNullException(nameof(openStoryForcesDialog));
            if (openStoryDriftsDialog == null) throw new ArgumentNullException(nameof(openStoryDriftsDialog));
            if (openStoryMaxOverAverageDisplacementsDialog == null) throw new ArgumentNullException(nameof(openStoryMaxOverAverageDisplacementsDialog));
            if (openStoryMaxOverAverageDriftsDialog == null) throw new ArgumentNullException(nameof(openStoryMaxOverAverageDriftsDialog));
            if (openMassSummaryByStoryDialog == null) throw new ArgumentNullException(nameof(openMassSummaryByStoryDialog));

            AnalysisResultTables = new ObservableCollection<AnalysisResultItem>();
            EtabsTableItems = new ObservableCollection<object>();
            ExportAnalysisResultTableCommand = new RelayCommand<AnalysisResultItem>(
                runAnalysisResult,
                item => item != null && _canUseActiveModel());
            ExportEtabsTableItemCommand = new RelayCommand<object>(
                runEtabsTableItem,
                item => item != null && _canUseActiveModel());
            GetBaseReactionsCommand = new RelayCommand(openGetBaseReactionsDialog, _canExecuteEtabsAction);
            GetModalMassParticipationRatiosCommand = new RelayCommand(openModalMassParticipationRatiosDialog, _canExecuteEtabsAction);
            GetStoryForcesCommand = new RelayCommand(openStoryForcesDialog, _canExecuteEtabsAction);
            GetStoryDriftsCommand = new RelayCommand(openStoryDriftsDialog, _canExecuteEtabsAction);
            GetStoryMaxOverAverageDisplacementsCommand = new RelayCommand(openStoryMaxOverAverageDisplacementsDialog, _canExecuteEtabsAction);
            GetStoryMaxOverAverageDriftsCommand = new RelayCommand(openStoryMaxOverAverageDriftsDialog, _canExecuteEtabsAction);
            GetMassSummaryByStoryCommand = new RelayCommand(openMassSummaryByStoryDialog, _canExecuteEtabsAction);
        }

        public ObservableCollection<AnalysisResultItem> AnalysisResultTables { get; private set; }

        public ObservableCollection<object> EtabsTableItems { get; private set; }

        public ICommand ExportAnalysisResultTableCommand { get; private set; }

        public ICommand ExportEtabsTableItemCommand { get; private set; }

        public ICommand GetBaseReactionsCommand { get; private set; }

        public ICommand GetModalMassParticipationRatiosCommand { get; private set; }

        public ICommand GetStoryForcesCommand { get; private set; }

        public ICommand GetStoryDriftsCommand { get; private set; }

        public ICommand GetStoryMaxOverAverageDisplacementsCommand { get; private set; }

        public ICommand GetStoryMaxOverAverageDriftsCommand { get; private set; }

        public ICommand GetMassSummaryByStoryCommand { get; private set; }

        public string ActiveTableCategory
        {
            get
            {
                return string.IsNullOrWhiteSpace(_activeTableCategory)
                    ? "ANALYSIS RESULTS"
                    : _activeTableCategory;
            }
            set
            {
                if (_activeTableCategory == value)
                {
                    return;
                }

                _activeTableCategory = value;
                OnPropertyChanged();
            }
        }

        public string ActiveAnalysisResultsGroup
        {
            get { return _activeAnalysisResultsGroup; }
            set
            {
                if (_activeAnalysisResultsGroup == value)
                {
                    return;
                }

                _activeAnalysisResultsGroup = value;
                OnPropertyChanged();
            }
        }

        public string SelectedAnalysisResultTable
        {
            get { return _selectedAnalysisResultTable; }
            set
            {
                if (_selectedAnalysisResultTable == value)
                {
                    return;
                }

                _selectedAnalysisResultTable = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AnalysisResultPlaceholderText));
            }
        }

        public string AnalysisResultPlaceholderText
        {
            get
            {
                return string.IsNullOrWhiteSpace(SelectedAnalysisResultTable)
                    ? "Select an ETABS result table from the tree."
                    : SelectedAnalysisResultTable;
            }
        }

        public void SetTableGroup(string category, string groupName)
        {
            ActiveTableCategory = string.IsNullOrWhiteSpace(category)
                ? "ANALYSIS RESULTS"
                : category;

            ActiveAnalysisResultsGroup = string.IsNullOrWhiteSpace(groupName)
                ? "Base Reactions"
                : groupName;

            EtabsTableItems.Clear();
            AnalysisResultTables.Clear();

            if (string.Equals(ActiveTableCategory, "Element Manipulation", StringComparison.OrdinalIgnoreCase))
            {
                SetElementConnectivityGroup(groupName);
                return;
            }

            if (string.Equals(ActiveTableCategory, "MISCELLANEOUS DATA", StringComparison.OrdinalIgnoreCase))
            {
                SetMiscellaneousDataGroup(groupName);
                return;
            }

            AnalysisResultGroup group = EtabsAnalysisResultRegistry.CreateGroupForNavigation(ActiveAnalysisResultsGroup);
            ActiveAnalysisResultsGroup = group.Name;
            foreach (AnalysisResultItem item in group.Items)
            {
                AnalysisResultTables.Add(item);
                EtabsTableItems.Add(item);
            }

            AnalysisResultItem matchingItem = null;
            foreach (AnalysisResultItem item in AnalysisResultTables)
            {
                if (string.Equals(item.Title, groupName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.EtabsTableName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    matchingItem = item;
                    break;
                }
            }

            SelectedAnalysisResultTable = matchingItem == null
                ? (AnalysisResultTables.Count > 0 ? AnalysisResultTables[0].Title : null)
                : matchingItem.Title;
        }

        private void SetElementConnectivityGroup(string groupName)
        {
            ElementConnectivityGroup group = EtabsElementConnectivityRegistry.CreateGroupForNavigation(groupName);
            ActiveAnalysisResultsGroup = group.Name;
            foreach (ElementConnectivityItem item in group.Items)
            {
                EtabsTableItems.Add(item);
            }

            ElementConnectivityItem matchingItem = null;
            foreach (ElementConnectivityItem item in group.Items)
            {
                if (string.Equals(item.Title, groupName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.EtabsTableName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    matchingItem = item;
                    break;
                }
            }

            SelectedAnalysisResultTable = matchingItem == null
                ? (group.Items.Count > 0 ? group.Items[0].Title : null)
                : matchingItem.Title;
        }

        private void SetMiscellaneousDataGroup(string groupName)
        {
            MiscellaneousDataGroup group = EtabsMiscellaneousDataRegistry.CreateGroupForNavigation(groupName);
            ActiveAnalysisResultsGroup = group.Name;
            foreach (MiscellaneousDataItem item in group.Items)
            {
                EtabsTableItems.Add(item);
            }

            MiscellaneousDataItem matchingItem = null;
            foreach (MiscellaneousDataItem item in group.Items)
            {
                if (string.Equals(item.Title, groupName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(item.EtabsTableName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    matchingItem = item;
                    break;
                }
            }

            SelectedAnalysisResultTable = matchingItem == null
                ? (group.Items.Count > 0 ? group.Items[0].Title : null)
                : matchingItem.Title;
        }
    }
}
