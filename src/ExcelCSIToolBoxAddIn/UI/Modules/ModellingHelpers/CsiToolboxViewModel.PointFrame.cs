using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private void SelectPointsByUniqueName()
        {
            ShowOperationResult(_useCases.SelectPointsByUniqueName.Execute());
        }

        private void SelectFramesByUniqueName()
        {
            ShowOperationResult(_useCases.SelectFramesByUniqueName.Execute());
        }

        private void AddPointByCartesian()
        {
            ShowOperationResult(_useCases.AddPointsByCartesian.Execute());
        }

        private void AddFramesByCoordinates()
        {
            ShowOperationResult(_useCases.AddFramesByCoordinates.Execute());
        }

        private void AddFramesByPointNames()
        {
            ShowOperationResult(_useCases.AddFramesByPointNames.Execute());
        }

        private void GetSelectedPoints()
        {
            ShowOperationResult(_useCases.GetSelectedPoints.Execute());
        }

        private void GetSelectedFrames()
        {
            ShowOperationResult(_useCases.GetSelectedFrames.Execute());
        }

        private void CreateShellAreasFromSelectedFrames()
        {
            if (!TryPickShellBoundaryFramesInteractively())
            {
                return;
            }

            var propertyName = PromptForShellPropertyName();
            if (propertyName == null)
            {
                return;
            }

            ShowOperationResult(_useCases.CreateShellAreasFromSelectedFrames.Execute(propertyName));
        }

        private bool TryPickShellBoundaryFramesInteractively()
        {
            OperationResult clearResult = _csiConnectionService.ClearSelection();
            if (!clearResult.IsSuccess)
            {
                ShowWarning(clearResult.Message);
                return false;
            }

            var window = new InteractiveSelectionWindow(
                "Create Shell Area",
                "Select the frame boundary objects in ETABS, then click Create Shell.",
                "Waiting for at least three frame objects...",
                "Select frame objects only.",
                "Selected objects must be frame objects.",
                "Frame",
                () => _csiConnectionService.GetSelectedObjectsFromActiveModel(),
                true,
                "Create Shell",
                3,
                "{0} frame object(s) selected. Click Create Shell to continue.",
                "Only frame objects can be used for shell area creation.");

            Window owner = GetActiveOwnerWindow();
            if (owner != null && !ReferenceEquals(owner, window))
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            ActivateConnectedCsiWindow();
            bool? result = window.ShowDialog();
            return result == true &&
                   window.SelectedObjects != null &&
                   window.SelectedObjects.Count >= 3 &&
                   AllSelectedObjectsAreFrames(window.SelectedObjects);
        }

        private static bool AllSelectedObjectsAreFrames(IReadOnlyList<CsiSelectedObjectDto> selectedObjects)
        {
            if (selectedObjects == null || selectedObjects.Count == 0)
            {
                return false;
            }

            foreach (CsiSelectedObjectDto selectedObject in selectedObjects)
            {
                if (selectedObject == null ||
                    !string.Equals(selectedObject.ObjectType, "Frame", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private static string PromptForShellPropertyName()
        {
            var dialog = new Window
            {
                Title = "Shell Property",
                Width = 360,
                Height = 150,
                MinWidth = 360,
                MinHeight = 150,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 12
            };

            var root = new StackPanel { Margin = new Thickness(14) };
            var label = new TextBlock
            {
                Text = "Enter shell property name. Leave blank to use Default.",
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            var textBox = new TextBox
            {
                Text = "Default",
                Height = 26,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 12)
            };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var okButton = new Button
            {
                Content = "OK",
                Width = 74,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true
            };
            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 74,
                IsCancel = true
            };

            okButton.Click += delegate
            {
                dialog.DialogResult = true;
                dialog.Close();
            };
            cancelButton.Click += delegate
            {
                dialog.DialogResult = false;
                dialog.Close();
            };

            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            root.Children.Add(label);
            root.Children.Add(textBox);
            root.Children.Add(buttons);
            dialog.Content = root;

            var result = dialog.ShowDialog();
            return result == true ? textBox.Text : null;
        }
    }
}
