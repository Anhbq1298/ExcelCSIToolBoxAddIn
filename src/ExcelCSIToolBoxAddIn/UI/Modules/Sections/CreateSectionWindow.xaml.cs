using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class CreateSectionWindow : Window
    {
        public CreateSectionWindow(CsiToolboxViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        private void IshapeButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreateIshapeSectionCommand);
        }

        private void ChannelButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreateChannelSectionCommand);
        }

        private void AngleButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreateAngleSectionCommand);
        }

        private void TubeButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreateTubeSectionCommand);
        }

        private void PipeButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreatePipeSectionCommand);
        }

        private void RectangleButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreateConcreteRectangleSectionCommand);
        }

        private void CircleButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommandAndClose(vm => vm.CreateConcreteCircleSectionCommand);
        }

        private void ExecuteCommandAndClose(Func<CsiToolboxViewModel, System.Windows.Input.ICommand> commandSelector)
        {
            if (DataContext is CsiToolboxViewModel vm)
            {
                var command = commandSelector(vm);
                if (command != null && command.CanExecute(null))
                {
                    this.DialogResult = true;
                    this.Close();
                    command.Execute(null);
                }
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
