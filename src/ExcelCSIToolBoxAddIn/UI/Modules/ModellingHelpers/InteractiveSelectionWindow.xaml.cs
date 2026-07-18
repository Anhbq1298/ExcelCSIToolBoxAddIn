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
        private readonly bool _requiresManualConfirmation;
        private readonly int _minimumMatchingCount;
        private readonly string _readyMessageFormat;
        private readonly string _mixedSelectionMessage;
        private readonly bool _ignoreNonMatchingObjects;
        private readonly DispatcherTimer _timer;

        public InteractiveSelectionWindow(
            string title,
            string instruction,
            string waitingMessage,
            string multipleSelectionMessage,
            string wrongTypeMessage,
            string requiredObjectType,
            Func<OperationResult<IReadOnlyList<CsiSelectedObjectDto>>> readSelection,
            bool requiresManualConfirmation = false,
            string confirmButtonText = null,
            int minimumMatchingCount = 1,
            string readyMessageFormat = null,
            string mixedSelectionMessage = null,
            bool ignoreNonMatchingObjects = false)
        {
            InitializeComponent();

            Title = string.IsNullOrWhiteSpace(title) ? "Pick Object" : title;
            InstructionTextBlock.Text = instruction;
            StatusTextBlock.Text = waitingMessage;
            ConfirmButton.Visibility = requiresManualConfirmation ? Visibility.Visible : Visibility.Collapsed;
            ConfirmButton.IsEnabled = false;
            if (!string.IsNullOrWhiteSpace(confirmButtonText))
            {
                ConfirmButton.Content = confirmButtonText;
            }

            _waitingMessage = waitingMessage;
            _multipleSelectionMessage = multipleSelectionMessage;
            _wrongTypeMessage = wrongTypeMessage;
            _requiredObjectType = requiredObjectType;
            _readSelection = readSelection ?? throw new ArgumentNullException(nameof(readSelection));
            _requiresManualConfirmation = requiresManualConfirmation;
            _minimumMatchingCount = Math.Max(1, minimumMatchingCount);
            _readyMessageFormat = string.IsNullOrWhiteSpace(readyMessageFormat)
                ? "{0} object(s) selected."
                : readyMessageFormat;
            _mixedSelectionMessage = string.IsNullOrWhiteSpace(mixedSelectionMessage)
                ? wrongTypeMessage
                : mixedSelectionMessage;
            _ignoreNonMatchingObjects = ignoreNonMatchingObjects;

            _timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(350)
            };
            _timer.Tick += Timer_Tick;

            Loaded += delegate { _timer.Start(); };
            Closed += delegate { _timer.Stop(); };
        }

        public CsiSelectedObjectDto SelectedObject { get; private set; }

        public IReadOnlyList<CsiSelectedObjectDto> SelectedObjects { get; private set; }

        private void Timer_Tick(object sender, EventArgs e)
        {
            OperationResult<IReadOnlyList<CsiSelectedObjectDto>> result = _readSelection();
            if (!result.IsSuccess)
            {
                SelectedObjects = null;
                ConfirmButton.IsEnabled = false;
                StatusTextBlock.Text = string.IsNullOrWhiteSpace(result.Message)
                    ? _waitingMessage
                    : result.Message;
                return;
            }

            IReadOnlyList<CsiSelectedObjectDto> selectedObjects = result.Data;
            if (selectedObjects == null || selectedObjects.Count == 0)
            {
                SelectedObjects = null;
                ConfirmButton.IsEnabled = false;
                StatusTextBlock.Text = _waitingMessage;
                return;
            }

            int matchingCount = 0;
            int nonMatchingCount = 0;
            CsiSelectedObjectDto matchingObject = null;
            var matchingObjects = new List<CsiSelectedObjectDto>();
            foreach (CsiSelectedObjectDto selectedObject in selectedObjects)
            {
                if (selectedObject == null)
                {
                    nonMatchingCount++;
                    continue;
                }

                if (string.Equals(selectedObject.ObjectType, _requiredObjectType, StringComparison.OrdinalIgnoreCase))
                {
                    matchingCount++;
                    matchingObject = selectedObject;
                    matchingObjects.Add(selectedObject);
                    continue;
                }

                nonMatchingCount++;
            }

            if (_requiresManualConfirmation)
            {
                UpdateManualConfirmationState(matchingObjects, matchingCount, nonMatchingCount);
                return;
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

        private void UpdateManualConfirmationState(
            IReadOnlyList<CsiSelectedObjectDto> matchingObjects,
            int matchingCount,
            int nonMatchingCount)
        {
            if (nonMatchingCount > 0 && !_ignoreNonMatchingObjects)
            {
                SelectedObjects = null;
                ConfirmButton.IsEnabled = false;
                StatusTextBlock.Text = _mixedSelectionMessage;
                return;
            }

            if (matchingCount < _minimumMatchingCount)
            {
                SelectedObjects = null;
                ConfirmButton.IsEnabled = false;
                StatusTextBlock.Text = _waitingMessage;
                return;
            }

            SelectedObjects = matchingObjects;
            ConfirmButton.IsEnabled = true;
            StatusTextBlock.Text = string.Format(_readyMessageFormat, matchingCount) +
                                   (nonMatchingCount > 0
                                       ? " " + nonMatchingCount + " non-matching object(s) will be ignored."
                                       : string.Empty);
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_requiresManualConfirmation ||
                SelectedObjects == null ||
                SelectedObjects.Count < _minimumMatchingCount)
            {
                return;
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
