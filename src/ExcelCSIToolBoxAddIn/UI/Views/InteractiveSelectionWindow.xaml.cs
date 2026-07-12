using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Threading;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    public partial class InteractiveSelectionWindow : Window
    {
        private readonly string _waitingMessage;
        private readonly string _multipleSelectionMessage;
        private readonly string _wrongTypeMessage;
        private readonly string _requiredObjectType;
        private readonly Func<OperationResult<IReadOnlyList<CsiSelectedObjectDto>>> _readSelection;
        private readonly DispatcherTimer _timer;

        public InteractiveSelectionWindow(
            string title,
            string instruction,
            string waitingMessage,
            string multipleSelectionMessage,
            string wrongTypeMessage,
            string requiredObjectType,
            Func<OperationResult<IReadOnlyList<CsiSelectedObjectDto>>> readSelection)
        {
            InitializeComponent();

            Title = string.IsNullOrWhiteSpace(title) ? "Pick Object" : title;
            InstructionTextBlock.Text = instruction;
            StatusTextBlock.Text = waitingMessage;

            _waitingMessage = waitingMessage;
            _multipleSelectionMessage = multipleSelectionMessage;
            _wrongTypeMessage = wrongTypeMessage;
            _requiredObjectType = requiredObjectType;
            _readSelection = readSelection ?? throw new ArgumentNullException(nameof(readSelection));

            _timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(350)
            };
            _timer.Tick += Timer_Tick;

            Loaded += delegate { _timer.Start(); };
            Closed += delegate { _timer.Stop(); };
        }

        public CsiSelectedObjectDto SelectedObject { get; private set; }

        private void Timer_Tick(object sender, EventArgs e)
        {
            OperationResult<IReadOnlyList<CsiSelectedObjectDto>> result = _readSelection();
            if (!result.IsSuccess)
            {
                StatusTextBlock.Text = string.IsNullOrWhiteSpace(result.Message)
                    ? _waitingMessage
                    : result.Message;
                return;
            }

            IReadOnlyList<CsiSelectedObjectDto> selectedObjects = result.Data;
            if (selectedObjects == null || selectedObjects.Count == 0)
            {
                StatusTextBlock.Text = _waitingMessage;
                return;
            }

            int matchingCount = 0;
            CsiSelectedObjectDto matchingObject = null;
            foreach (CsiSelectedObjectDto selectedObject in selectedObjects)
            {
                if (selectedObject == null)
                {
                    continue;
                }

                if (string.Equals(selectedObject.ObjectType, _requiredObjectType, StringComparison.OrdinalIgnoreCase))
                {
                    matchingCount++;
                    matchingObject = selectedObject;
                }
            }

            if (matchingCount == 1)
            {
                SelectedObject = matchingObject;
                DialogResult = true;
                Close();
                return;
            }

            StatusTextBlock.Text = matchingCount > 1 ? _multipleSelectionMessage : _wrongTypeMessage;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
