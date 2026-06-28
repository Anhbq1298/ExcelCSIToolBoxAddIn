using System;
using System.Windows;
using System.Windows.Controls;

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
            var propertyName = PromptForShellPropertyName();
            if (propertyName == null)
            {
                return;
            }

            ShowOperationResult(_useCases.CreateShellAreasFromSelectedFrames.Execute(propertyName));
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
