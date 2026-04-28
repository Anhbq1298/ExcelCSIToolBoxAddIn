using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using ExcelCSIToolBoxAddIn.UI.ViewModels;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    internal sealed class AiAgentTaskPaneHost : UserControl
    {
        private static readonly Color Ink = Color.FromArgb(38, 43, 50);
        private static readonly Color MutedInk = Color.FromArgb(110, 118, 128);
        private static readonly Color Canvas = Color.FromArgb(246, 242, 234);
        private static readonly Color Card = Color.FromArgb(255, 253, 248);
        private static readonly Color Line = Color.FromArgb(224, 216, 204);
        private static readonly Color Accent = Color.FromArgb(38, 94, 116);
        private static readonly Color AccentDisabled = Color.FromArgb(176, 196, 201);

        private readonly AiAgentChatViewModel _viewModel;
        private readonly RichTextBox _chatBox;
        private readonly RichTextBox _traceBox;
        private readonly TextBox _inputBox;
        private readonly Button _sendButton;
        private readonly Button _clearButton;
        private readonly Label _statusLabel;
        private bool _syncingInput;

        public AiAgentTaskPaneHost()
        {
            Dock = DockStyle.Fill;
            BackColor = Canvas;

            _viewModel = new AiAgentChatViewModel();
            _chatBox = CreateRichTextBox(new Font("Segoe UI", 9.5F));
            _traceBox = CreateRichTextBox(new Font("Consolas", 8.5F));
            _inputBox = new TextBox
            {
                AcceptsReturn = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5F),
                ForeColor = Ink,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            _sendButton = CreateButton("Send", Accent, Color.White);
            _clearButton = CreateButton("Clear", Color.FromArgb(239, 234, 226), Ink);
            _statusLabel = new Label
            {
                AutoEllipsis = true,
                BackColor = Color.Transparent,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = MutedInk,
                TextAlign = ContentAlignment.MiddleLeft
            };

            BuildLayout();
            WireEvents();
            UpdateAll();
        }

        private static RichTextBox CreateRichTextBox(Font font)
        {
            return new RichTextBox
            {
                BackColor = Card,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = font,
                ForeColor = Ink,
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
        }

        private static Button CreateButton(string text, Color backColor, Color foreColor)
        {
            var button = new Button
            {
                BackColor = backColor,
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = foreColor,
                Margin = new Padding(0, 0, 0, 6),
                Text = text,
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                BackColor = Canvas,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(14),
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 132F));

            var header = new Panel { BackColor = Canvas, Dock = DockStyle.Fill };
            var title = new Label
            {
                Dock = DockStyle.Top,
                Font = new Font("Georgia", 16F, FontStyle.Bold),
                ForeColor = Ink,
                Height = 31,
                Text = "AI Agent Studio",
                TextAlign = ContentAlignment.MiddleLeft
            };
            var subtitle = new Label
            {
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 8.75F),
                ForeColor = MutedInk,
                Height = 22,
                Text = "Elegant pane build 2026.04.29. Ctrl+Enter sends.",
                TextAlign = ContentAlignment.MiddleLeft
            };
            header.Controls.Add(subtitle);
            header.Controls.Add(title);

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(CreateCard("Conversation", _chatBox), 0, 1);
            root.Controls.Add(new Panel { BackColor = Canvas, Dock = DockStyle.Fill }, 0, 2);
            root.Controls.Add(CreateCard("Tool Trace", _traceBox), 0, 3);
            root.Controls.Add(CreateComposer(), 0, 4);
            Controls.Add(root);
        }

        private Control CreateCard(string title, Control content)
        {
            var card = new BorderedPanel
            {
                BackColor = Card,
                BorderColor = Line,
                Dock = DockStyle.Fill,
                Padding = new Padding(12)
            };
            var label = new Label
            {
                BackColor = Color.Transparent,
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI Semibold", 8.75F, FontStyle.Bold),
                ForeColor = MutedInk,
                Height = 22,
                Text = title.ToUpperInvariant()
            };
            card.Controls.Add(content);
            card.Controls.Add(label);
            content.BringToFront();
            return card;
        }

        private Control CreateComposer()
        {
            var outer = new TableLayoutPanel
            {
                BackColor = Canvas,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                RowCount = 3
            };
            outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 10F));
            outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));

            var inputRow = new TableLayoutPanel
            {
                BackColor = Canvas,
                ColumnCount = 2,
                Dock = DockStyle.Fill
            };
            inputRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inputRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92F));

            var inputCard = new BorderedPanel
            {
                BackColor = Color.White,
                BorderColor = Line,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 8, 10, 8)
            };
            inputCard.Controls.Add(_inputBox);

            var buttonPanel = new TableLayoutPanel
            {
                BackColor = Canvas,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 0, 0, 0),
                RowCount = 2
            };
            buttonPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            buttonPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            buttonPanel.Controls.Add(_sendButton, 0, 0);
            buttonPanel.Controls.Add(_clearButton, 0, 1);

            var statusPanel = new Panel { BackColor = Canvas, Dock = DockStyle.Fill };
            statusPanel.Controls.Add(_statusLabel);
            statusPanel.Controls.Add(new StatusDotPanel { BackColor = Canvas, Dock = DockStyle.Left, DotColor = Accent, Width = 18 });
            _statusLabel.BringToFront();

            inputRow.Controls.Add(inputCard, 0, 0);
            inputRow.Controls.Add(buttonPanel, 1, 0);
            outer.Controls.Add(new Panel { BackColor = Canvas, Dock = DockStyle.Fill }, 0, 0);
            outer.Controls.Add(inputRow, 0, 1);
            outer.Controls.Add(statusPanel, 0, 2);
            return outer;
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
            if (!_syncingInput)
            {
                _viewModel.UserInput = _inputBox.Text;
            }
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
            _statusLabel.Text = _viewModel.IsBusy ? "Thinking..." : _viewModel.StatusText;
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
            _chatBox.Clear();
            if (_viewModel.Messages.Count == 0)
            {
                AppendMuted(_chatBox, "Start a conversation with the model.\nTip: ask about selected frames, section properties, or model units.");
                ScrollToEnd(_chatBox);
                return;
            }

            foreach (AiAgentChatMessageViewModel message in _viewModel.Messages)
            {
                bool isUser = string.Equals(message.Role, "User", StringComparison.OrdinalIgnoreCase);
                AppendRole(_chatBox, isUser ? "You" : "Assistant", isUser ? Accent : Color.FromArgb(122, 77, 45));
                AppendBody(_chatBox, message.Content);
                _chatBox.AppendText(Environment.NewLine);
            }
            ScrollToEnd(_chatBox);
        }

        private void UpdateTrace()
        {
            _traceBox.Clear();
            AiAgentToolTraceViewModel trace = _viewModel.LastToolTrace;
            if (!trace.ToolWasCalled)
            {
                AppendMuted(_traceBox, "No tool call for the last response.");
                return;
            }

            AppendTraceLine("Tool", trace.ToolName);
            AppendTraceLine("Arguments", trace.ToolArguments);
            AppendTraceLine("Succeeded", trace.ToolSucceeded ? "Yes" : "No");
            AppendTraceLine("Message", trace.ToolMessage);
            AppendTraceLine("Result", trace.ToolResultJson);
            ScrollToEnd(_traceBox);
        }

        private void AppendTraceLine(string label, string value)
        {
            AppendRole(_traceBox, label, Accent);
            AppendBody(_traceBox, string.IsNullOrWhiteSpace(value) ? "(none)" : value);
        }

        private static void AppendRole(RichTextBox box, string role, Color color)
        {
            box.SelectionColor = color;
            box.SelectionFont = new Font(box.Font, FontStyle.Bold);
            box.AppendText(role + Environment.NewLine);
            box.SelectionFont = box.Font;
            box.SelectionColor = Ink;
        }

        private static void AppendBody(RichTextBox box, string text)
        {
            box.SelectionColor = Ink;
            box.SelectionFont = box.Font;
            box.AppendText((text ?? string.Empty).Trim());
            box.AppendText(Environment.NewLine);
        }

        private static void AppendMuted(RichTextBox box, string text)
        {
            box.SelectionColor = MutedInk;
            box.SelectionFont = new Font(box.Font, FontStyle.Italic);
            box.AppendText(text);
            box.SelectionFont = box.Font;
            box.SelectionColor = Ink;
        }

        private void UpdateButtons()
        {
            bool canSend = _viewModel.SendCommand.CanExecute(null);
            _sendButton.Enabled = canSend;
            _sendButton.BackColor = canSend ? Accent : AccentDisabled;
            _clearButton.Enabled = _viewModel.ClearCommand.CanExecute(null);
            _inputBox.Enabled = !_viewModel.IsBusy;
        }

        private static void ScrollToEnd(RichTextBox box)
        {
            box.SelectionStart = box.TextLength;
            box.ScrollToCaret();
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

        private sealed class BorderedPanel : Panel
        {
            public Color BorderColor { get; set; }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(BorderColor))
                {
                    Rectangle rect = ClientRectangle;
                    rect.Width -= 1;
                    rect.Height -= 1;
                    e.Graphics.DrawRectangle(pen, rect);
                }
            }
        }

        private sealed class StatusDotPanel : Panel
        {
            public Color DotColor { get; set; }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(DotColor))
                {
                    e.Graphics.FillEllipse(brush, 4, 9, 8, 8);
                }
            }
        }
    }
}
