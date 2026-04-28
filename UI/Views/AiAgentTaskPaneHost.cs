using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    internal sealed class AiAgentTaskPaneHost : UserControl
    {
        private readonly AiAgentChatViewModel _viewModel;
        private readonly TextBox _chatTranscript;
        private readonly TextBox _toolTrace;
        private readonly TextBox _inputBox;
        private readonly Button _sendButton;
        private readonly Button _clearButton;
        private readonly Label _statusLabel;
        private bool _syncingInput;

        public AiAgentTaskPaneHost()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(248, 250, 252);

            _viewModel = new AiAgentChatViewModel();
            _chatTranscript = CreateReadOnlyTextBox();
            _toolTrace = CreateReadOnlyTextBox();
            _inputBox = new TextBox
            {
                AcceptsReturn = true,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            _sendButton = new Button
            {
                Dock = DockStyle.Fill,
                Text = "Send",
                UseVisualStyleBackColor = true
            };
            _clearButton = new Button
            {
                Dock = DockStyle.Fill,
                Text = "Clear",
                UseVisualStyleBackColor = true
            };
            _statusLabel = new Label
            {
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.5F),
                TextAlign = ContentAlignment.MiddleLeft
            };

            BuildLayout();
            WireEvents();
            UpdateAll();
        }

        private static TextBox CreateReadOnlyTextBox()
        {
            return new TextBox
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 92F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));

            var title = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                Text = "AI Agent",
                TextAlign = ContentAlignment.MiddleLeft
            };

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 320
            };
            split.Panel1.Controls.Add(_chatTranscript);
            split.Panel2.Controls.Add(_toolTrace);

            var inputPanel = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                RowCount = 1
            };
            inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88F));
            inputPanel.Controls.Add(_inputBox, 0, 0);

            var buttonPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2
            };
            buttonPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            buttonPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            buttonPanel.Controls.Add(_sendButton, 0, 0);
            buttonPanel.Controls.Add(_clearButton, 0, 1);
            inputPanel.Controls.Add(buttonPanel, 1, 0);

            root.Controls.Add(title, 0, 0);
            root.Controls.Add(split, 0, 1);
            root.Controls.Add(inputPanel, 0, 2);
            root.Controls.Add(_statusLabel, 0, 3);
            Controls.Add(root);
        }

        private void WireEvents()
        {
            _viewModel.Messages.CollectionChanged += Messages_CollectionChanged;
            _viewModel.PropertyChanged += ViewModel_PropertyChanged;
            _viewModel.LastToolTrace.PropertyChanged += LastToolTrace_PropertyChanged;
            _viewModel.SendCommand.CanExecuteChanged += Command_CanExecuteChanged;

            _inputBox.TextChanged += InputBox_TextChanged;
            _inputBox.KeyDown += InputBox_KeyDown;
            _sendButton.Click += SendButton_Click;
            _clearButton.Click += ClearButton_Click;
        }

        private void Messages_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            RunOnUiThread(UpdateTranscript);
        }

        private void ViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            RunOnUiThread(UpdateAll);
        }

        private void LastToolTrace_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            RunOnUiThread(UpdateTrace);
        }

        private void Command_CanExecuteChanged(object sender, EventArgs e)
        {
            RunOnUiThread(UpdateButtons);
        }

        private void InputBox_TextChanged(object sender, EventArgs e)
        {
            if (_syncingInput)
            {
                return;
            }

            _viewModel.UserInput = _inputBox.Text;
        }

        private void InputBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && e.Control)
            {
                TrySend();
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            TrySend();
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            if (_viewModel.ClearCommand.CanExecute(null))
            {
                _viewModel.ClearCommand.Execute(null);
            }
        }

        private void TrySend()
        {
            _viewModel.UserInput = _inputBox.Text;
            if (_viewModel.SendCommand.CanExecute(null))
            {
                _viewModel.SendCommand.Execute(null);
            }
        }

        private void UpdateAll()
        {
            SyncInputFromViewModel();
            UpdateTranscript();
            UpdateTrace();
            UpdateButtons();
            _statusLabel.Text = _viewModel.StatusText;
        }

        private void SyncInputFromViewModel()
        {
            if (_inputBox.Text == _viewModel.UserInput)
            {
                return;
            }

            _syncingInput = true;
            _inputBox.Text = _viewModel.UserInput ?? string.Empty;
            _inputBox.SelectionStart = _inputBox.TextLength;
            _syncingInput = false;
        }

        private void UpdateTranscript()
        {
            var builder = new StringBuilder();
            foreach (AiAgentChatMessageViewModel message in _viewModel.Messages)
            {
                builder.Append(message.Role);
                builder.AppendLine(":");
                builder.AppendLine(message.Content);
                builder.AppendLine();
            }

            _chatTranscript.Text = builder.Length == 0
                ? "Type a question below, then press Send or Ctrl+Enter."
                : builder.ToString();
            ScrollToEnd(_chatTranscript);
        }

        private void UpdateTrace()
        {
            AiAgentToolTraceViewModel trace = _viewModel.LastToolTrace;
            if (!trace.ToolWasCalled)
            {
                _toolTrace.Text = "No tool call for the last response.";
                return;
            }

            var builder = new StringBuilder();
            builder.AppendLine("Tool:");
            builder.AppendLine(trace.ToolName);
            builder.AppendLine();
            builder.AppendLine("Arguments:");
            builder.AppendLine(trace.ToolArguments);
            builder.AppendLine();
            builder.AppendLine("Succeeded:");
            builder.AppendLine(trace.ToolSucceeded ? "Yes" : "No");
            builder.AppendLine();
            builder.AppendLine("Message:");
            builder.AppendLine(trace.ToolMessage);
            builder.AppendLine();
            builder.AppendLine("Result:");
            builder.AppendLine(trace.ToolResultJson);

            _toolTrace.Text = builder.ToString();
            ScrollToEnd(_toolTrace);
        }

        private void UpdateButtons()
        {
            _sendButton.Enabled = _viewModel.SendCommand.CanExecute(null);
            _clearButton.Enabled = _viewModel.ClearCommand.CanExecute(null);
            _inputBox.Enabled = !_viewModel.IsBusy;
        }

        private static void ScrollToEnd(TextBox textBox)
        {
            textBox.SelectionStart = textBox.TextLength;
            textBox.ScrollToCaret();
        }

        private void RunOnUiThread(Action action)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(action);
                return;
            }

            action();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _viewModel.Messages.CollectionChanged -= Messages_CollectionChanged;
                _viewModel.PropertyChanged -= ViewModel_PropertyChanged;
                _viewModel.LastToolTrace.PropertyChanged -= LastToolTrace_PropertyChanged;
                _viewModel.SendCommand.CanExecuteChanged -= Command_CanExecuteChanged;
            }

            base.Dispose(disposing);
        }
    }
}
