using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private string _frameSectionSearchText;
        private string _areaSectionSearchText;
        private string _frameModifierWarningText;
        private string _areaModifierWarningText;
        private bool _isBulkUpdatingStiffnessSelections;

        private void InitializeStiffnessModifierPage()
        {
            FrameStiffnessSections = new ObservableCollection<SelectableSectionItem>();
            AreaStiffnessSections = new ObservableCollection<SelectableSectionItem>();
            FrameModifierFields = CreateFrameModifierFields();
            AreaModifierFields = CreateAreaModifierFields();

            FrameStiffnessSectionView = CollectionViewSource.GetDefaultView(FrameStiffnessSections);
            FrameStiffnessSectionView.Filter = FilterFrameStiffnessSection;
            AreaStiffnessSectionView = CollectionViewSource.GetDefaultView(AreaStiffnessSections);
            AreaStiffnessSectionView.Filter = FilterAreaStiffnessSection;
        }

        public ObservableCollection<SelectableSectionItem> FrameStiffnessSections { get; private set; }
        public ObservableCollection<SelectableSectionItem> AreaStiffnessSections { get; private set; }
        public ObservableCollection<ModifierFieldViewModel> FrameModifierFields { get; private set; }
        public ObservableCollection<ModifierFieldViewModel> AreaModifierFields { get; private set; }
        public ICollectionView FrameStiffnessSectionView { get; private set; }
        public ICollectionView AreaStiffnessSectionView { get; private set; }

        public string FrameSectionSearchText
        {
            get { return _frameSectionSearchText; }
            set
            {
                if (_frameSectionSearchText == value)
                {
                    return;
                }

                _frameSectionSearchText = value;
                OnPropertyChanged();
                if (FrameStiffnessSectionView != null)
                {
                    FrameStiffnessSectionView.Refresh();
                }
            }
        }

        public string AreaSectionSearchText
        {
            get { return _areaSectionSearchText; }
            set
            {
                if (_areaSectionSearchText == value)
                {
                    return;
                }

                _areaSectionSearchText = value;
                OnPropertyChanged();
                if (AreaStiffnessSectionView != null)
                {
                    AreaStiffnessSectionView.Refresh();
                }
            }
        }

        public string FrameModifierWarningText
        {
            get { return _frameModifierWarningText; }
            private set
            {
                if (_frameModifierWarningText == value)
                {
                    return;
                }

                _frameModifierWarningText = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FrameModifierWarningVisibility));
            }
        }

        public string AreaModifierWarningText
        {
            get { return _areaModifierWarningText; }
            private set
            {
                if (_areaModifierWarningText == value)
                {
                    return;
                }

                _areaModifierWarningText = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AreaModifierWarningVisibility));
            }
        }

        public Visibility FrameModifierWarningVisibility
        {
            get { return string.IsNullOrWhiteSpace(FrameModifierWarningText) ? Visibility.Collapsed : Visibility.Visible; }
        }

        public Visibility AreaModifierWarningVisibility
        {
            get { return string.IsNullOrWhiteSpace(AreaModifierWarningText) ? Visibility.Collapsed : Visibility.Visible; }
        }

        private void RefreshFrameStiffnessSections()
        {
            if (!EnsureAttachedForStiffnessModifier())
            {
                return;
            }

            OperationResult<IReadOnlyList<string>> result = _csiConnectionService.GetFrameSectionNames();
            if (!result.IsSuccess)
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
                return;
            }

            FrameStiffnessSections.Clear();
            IEnumerable<string> names = result.Data ?? new List<string>();
            foreach (string name in names.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                SelectableSectionItem item = new SelectableSectionItem(name);
                item.PropertyChanged += FrameStiffnessSectionItem_PropertyChanged;
                FrameStiffnessSections.Add(item);
            }

            FrameModifierWarningText = string.Empty;
            ResetFrameModifierFields();
            StatusText = "Frame section list refreshed.";
        }

        private void RefreshAreaStiffnessSections()
        {
            if (!EnsureAttachedForStiffnessModifier())
            {
                return;
            }

            OperationResult<IReadOnlyList<string>> result = _csiConnectionService.GetAreaSectionNames();
            if (!result.IsSuccess)
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
                return;
            }

            AreaStiffnessSections.Clear();
            IEnumerable<string> names = result.Data ?? new List<string>();
            foreach (string name in names.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                SelectableSectionItem item = new SelectableSectionItem(name);
                item.PropertyChanged += AreaStiffnessSectionItem_PropertyChanged;
                AreaStiffnessSections.Add(item);
            }

            AreaModifierWarningText = string.Empty;
            ResetAreaModifierFields();
            StatusText = "Area section list refreshed.";
        }

        private void ApplyFrameStiffnessModifiers()
        {
            IReadOnlyList<SelectableSectionItem> selectedSections = GetSelectedSections(FrameStiffnessSections);
            if (!EnsureCanApplyModifiers(selectedSections))
            {
                return;
            }

            double[] modifiers;
            if (!TryReadModifierFields(FrameModifierFields, 8, out modifiers))
            {
                MessageBox.Show("Please enter valid numeric modifier values.", ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (SelectableSectionItem section in selectedSections)
            {
                OperationResult result = _csiConnectionService.SetFrameSectionModifiers(section.Name, modifiers);
                if (!result.IsSuccess)
                {
                    ShowOperationResult(OperationResult.Failure(result.Message));
                    return;
                }
            }

            MessageBox.Show(
                $"Frame stiffness modifiers applied to {selectedSections.Count} section(s).",
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            RefreshFrameModifierFieldsFromSelection();
        }

        private void ApplyAreaStiffnessModifiers()
        {
            IReadOnlyList<SelectableSectionItem> selectedSections = GetSelectedSections(AreaStiffnessSections);
            if (!EnsureCanApplyModifiers(selectedSections))
            {
                return;
            }

            double[] modifiers;
            if (!TryReadModifierFields(AreaModifierFields, 10, out modifiers))
            {
                MessageBox.Show("Please enter valid numeric modifier values.", ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (SelectableSectionItem section in selectedSections)
            {
                OperationResult result = _csiConnectionService.SetAreaSectionModifiers(section.Name, modifiers);
                if (!result.IsSuccess)
                {
                    ShowOperationResult(OperationResult.Failure(result.Message));
                    return;
                }
            }

            MessageBox.Show(
                $"Area stiffness modifiers applied to {selectedSections.Count} section(s).",
                ProductTitle,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            RefreshAreaModifierFieldsFromSelection();
        }

        private void SelectVisibleFrameStiffnessSections()
        {
            SetVisibleSectionSelection(FrameStiffnessSections, FrameSectionSearchText, true, RefreshFrameModifierFieldsFromSelection);
        }

        private void ClearFrameStiffnessSectionSelection()
        {
            SetSectionSelection(FrameStiffnessSections, false, RefreshFrameModifierFieldsFromSelection);
        }

        private void SelectVisibleAreaStiffnessSections()
        {
            SetVisibleSectionSelection(AreaStiffnessSections, AreaSectionSearchText, true, RefreshAreaModifierFieldsFromSelection);
        }

        private void ClearAreaStiffnessSectionSelection()
        {
            SetSectionSelection(AreaStiffnessSections, false, RefreshAreaModifierFieldsFromSelection);
        }

        private void ResetFrameModifierFields()
        {
            SetModifierFields(FrameModifierFields, CreateDefaultModifiers(8));
        }

        private void ResetAreaModifierFields()
        {
            SetModifierFields(AreaModifierFields, CreateDefaultModifiers(10));
        }

        private void RefreshFrameModifierFieldsFromSelection()
        {
            RefreshModifierFieldsFromSelection(
                FrameStiffnessSections,
                FrameModifierFields,
                8,
                _csiConnectionService.GetFrameSectionModifiers,
                warning => FrameModifierWarningText = warning);
        }

        private void RefreshAreaModifierFieldsFromSelection()
        {
            RefreshModifierFieldsFromSelection(
                AreaStiffnessSections,
                AreaModifierFields,
                10,
                _csiConnectionService.GetAreaSectionModifiers,
                warning => AreaModifierWarningText = warning);
        }

        private void RefreshModifierFieldsFromSelection(
            ObservableCollection<SelectableSectionItem> sections,
            ObservableCollection<ModifierFieldViewModel> fields,
            int expectedCount,
            Func<string, OperationResult<double[]>> getModifiers,
            Action<string> setWarning)
        {
            IReadOnlyList<SelectableSectionItem> selectedSections = GetSelectedSections(sections);
            if (selectedSections.Count == 0)
            {
                setWarning(string.Empty);
                return;
            }

            List<double[]> allModifiers = new List<double[]>();
            foreach (SelectableSectionItem section in selectedSections)
            {
                OperationResult<double[]> result = getModifiers(section.Name);
                if (!result.IsSuccess)
                {
                    setWarning(result.Message);
                    return;
                }

                allModifiers.Add(NormalizeModifierArray(result.Data, expectedCount));
            }

            SetModifierFields(fields, allModifiers[0]);
            bool sameValues = allModifiers.All(x => ModifierArraysMatch(allModifiers[0], x));
            setWarning(sameValues
                ? string.Empty
                : "Selected sections have different modifier values. Applying will overwrite all selected sections.");
        }

        private bool EnsureAttachedForStiffnessModifier()
        {
            OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelConnectionInfoDTO> result = _csiConnectionService.GetCurrentConnection();
            if (result.IsSuccess)
            {
                return true;
            }

            MessageBox.Show("Please attach to ETABS first.", ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private bool EnsureCanApplyModifiers(IReadOnlyList<SelectableSectionItem> selectedSections)
        {
            if (!EnsureAttachedForStiffnessModifier())
            {
                return false;
            }

            if (selectedSections == null || selectedSections.Count == 0)
            {
                MessageBox.Show("Please select at least one section.", ProductTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private bool FilterFrameStiffnessSection(object item)
        {
            SelectableSectionItem section = item as SelectableSectionItem;
            return SectionMatchesSearch(section, FrameSectionSearchText);
        }

        private bool FilterAreaStiffnessSection(object item)
        {
            SelectableSectionItem section = item as SelectableSectionItem;
            return SectionMatchesSearch(section, AreaSectionSearchText);
        }

        private static bool SectionMatchesSearch(SelectableSectionItem section, string searchText)
        {
            if (section == null)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(searchText))
            {
                return true;
            }

            return section.Name.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void FrameStiffnessSectionItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SelectableSectionItem.IsSelected))
            {
                if (_isBulkUpdatingStiffnessSelections)
                {
                    return;
                }

                RefreshFrameModifierFieldsFromSelection();
            }
        }

        private void AreaStiffnessSectionItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SelectableSectionItem.IsSelected))
            {
                if (_isBulkUpdatingStiffnessSelections)
                {
                    return;
                }

                RefreshAreaModifierFieldsFromSelection();
            }
        }

        private void SetVisibleSectionSelection(
            IEnumerable<SelectableSectionItem> sections,
            string searchText,
            bool isSelected,
            Action afterSelectionChanged)
        {
            if (sections == null)
            {
                return;
            }

            SetSectionSelection(
                sections.Where(section => SectionMatchesSearch(section, searchText)),
                isSelected,
                afterSelectionChanged);
        }

        private void SetSectionSelection(
            IEnumerable<SelectableSectionItem> sections,
            bool isSelected,
            Action afterSelectionChanged)
        {
            if (sections == null)
            {
                return;
            }

            try
            {
                _isBulkUpdatingStiffnessSelections = true;
                foreach (SelectableSectionItem section in sections.ToList())
                {
                    section.IsSelected = isSelected;
                }
            }
            finally
            {
                _isBulkUpdatingStiffnessSelections = false;
            }

            if (afterSelectionChanged != null)
            {
                afterSelectionChanged();
            }
        }

        private static IReadOnlyList<SelectableSectionItem> GetSelectedSections(IEnumerable<SelectableSectionItem> sections)
        {
            return sections == null
                ? new List<SelectableSectionItem>()
                : sections.Where(x => x.IsSelected).ToList();
        }

        private static bool TryReadModifierFields(
            ObservableCollection<ModifierFieldViewModel> fields,
            int expectedCount,
            out double[] modifiers)
        {
            modifiers = new double[expectedCount];
            if (fields == null || fields.Count != expectedCount)
            {
                return false;
            }

            for (int i = 0; i < expectedCount; i++)
            {
                double value;
                if (string.IsNullOrWhiteSpace(fields[i].ValueText) ||
                    !double.TryParse(fields[i].ValueText, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                    value < 0)
                {
                    return false;
                }

                modifiers[i] = value;
            }

            return true;
        }

        private static void SetModifierFields(ObservableCollection<ModifierFieldViewModel> fields, double[] modifiers)
        {
            if (fields == null || modifiers == null)
            {
                return;
            }

            for (int i = 0; i < fields.Count && i < modifiers.Length; i++)
            {
                fields[i].ValueText = modifiers[i].ToString("0.##########", CultureInfo.InvariantCulture);
            }
        }

        private static double[] NormalizeModifierArray(double[] source, int expectedCount)
        {
            double[] result = CreateDefaultModifiers(expectedCount);
            if (source == null)
            {
                return result;
            }

            for (int i = 0; i < source.Length && i < result.Length; i++)
            {
                result[i] = source[i];
            }

            return result;
        }

        private static bool ModifierArraysMatch(double[] left, double[] right)
        {
            if (left == null || right == null || left.Length != right.Length)
            {
                return false;
            }

            for (int i = 0; i < left.Length; i++)
            {
                if (Math.Abs(left[i] - right[i]) > 0.000000001)
                {
                    return false;
                }
            }

            return true;
        }

        private static double[] CreateDefaultModifiers(int count)
        {
            double[] modifiers = new double[count];
            for (int i = 0; i < modifiers.Length; i++)
            {
                modifiers[i] = 1;
            }

            return modifiers;
        }

        private static ObservableCollection<ModifierFieldViewModel> CreateFrameModifierFields()
        {
            return new ObservableCollection<ModifierFieldViewModel>
            {
                new ModifierFieldViewModel("Cross-section (axial) Area"),
                new ModifierFieldViewModel("Shear Area in 2 direction"),
                new ModifierFieldViewModel("Shear Area in 3 direction"),
                new ModifierFieldViewModel("Torsional Constant"),
                new ModifierFieldViewModel("Moment of Inertia about 2 axis"),
                new ModifierFieldViewModel("Moment of Inertia about 3 axis"),
                new ModifierFieldViewModel("Mass"),
                new ModifierFieldViewModel("Weight")
            };
        }

        private static ObservableCollection<ModifierFieldViewModel> CreateAreaModifierFields()
        {
            return new ObservableCollection<ModifierFieldViewModel>
            {
                new ModifierFieldViewModel("Membrane f11"),
                new ModifierFieldViewModel("Membrane f22"),
                new ModifierFieldViewModel("Membrane f12"),
                new ModifierFieldViewModel("Bending m11"),
                new ModifierFieldViewModel("Bending m22"),
                new ModifierFieldViewModel("Bending m12"),
                new ModifierFieldViewModel("Shear v13"),
                new ModifierFieldViewModel("Shear v23"),
                new ModifierFieldViewModel("Mass"),
                new ModifierFieldViewModel("Weight")
            };
        }
    }
}
