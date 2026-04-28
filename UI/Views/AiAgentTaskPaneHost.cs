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
        private static readonly Color MainBackground = Color.FromArgb(247, 251, 255);
        private static readonly Color ConversationBackground = Color.FromArgb(250, 253, 255);
        private static readonly Color LightBlueSurface = Color.FromArgb(234, 245, 252);
        private static readonly Color HeaderAccent = Color.FromArgb(31, 106, 165);
        private static readonly Color BorderBlue = Color.FromArgb(201, 221, 234);
        private static readonly Color AssistantBubble = Color.FromArgb(242, 247, 251);
        private static readonly Color Ink = Color.FromArgb(34, 34, 34);
        private static readonly Color MutedInk = Color.FromArgb(111, 127, 140);
        private static readonly Color ClearButtonBack = Color.FromArgb(234, 241, 246);
        private static readonly Color ClearButtonText = Color.FromArgb(46, 58, 68);

        private readonly AiAgentChatViewModel _viewModel;
        private readonly Label _subtitleLabel;
        private readonly Label _sap2000BadgeLabel;
        private readonly Label _etabsBadgeLabel;
        private readonly FlowLayoutPanel _conversationPanel;
        private readonly TextBox _inputBox;
        private readonly Button _sendButton;
        private readonly Button _clearButton;
        private readonly Label _statusLabel;
        private bool _syncingInput;

        public AiAgentTaskPaneHost()
        {
            Dock = DockStyle.Fill;
            BackColor = MainBackground;

            _viewModel = new AiAgentChatViewModel();
            _subtitleLabel = CreateHeaderSubtitle();
            _sap2000BadgeLabel = CreateBadgeLabel();
            _etabsBadgeLabel = CreateBadgeLabel();
            _conversationPanel = new FlowLayoutPanel
            {
                AutoScroll = true,
                BackColor = ConversationBackground,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(10),
                WrapContents = false
            };
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
            _sendButton = CreateButton("Send", HeaderAccent, Color.White);
            _clearButton = CreateButton("Clear", ClearButtonBack, ClearButtonText);
            _statusLabel = new Label
            {
                AutoEllipsis = true,
                BackColor = MainBackground,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = MutedInk,
                TextAlign = ContentAlignment.MiddleLeft
            };

            BuildLayout();
            WireEvents();
            UpdateAll();
        }

        private static Label CreateHeaderSubtitle()
        {
            return new Label
            {
                AutoEllipsis = true,
                BackColor = MainBackground,
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = MutedInk,
                Height = 20,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static Label CreateBadgeLabel()
        {
            return new Label
            {
                AutoSize = true,
                BackColor = LightBlueSurface,
                Font = new Font("Segoe UI Semibold", 8.25F, FontStyle.Bold),
                ForeColor = HeaderAccent,
                Margin = new Padding(0, 0, 6, 4),
                Padding = new Padding(8, 3, 8, 3)
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
                BackColor = MainBackground,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 86F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 94F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));

            root.Controls.Add(CreateHeader(), 0, 0);
            root.Controls.Add(CreateConversationCard(), 0, 1);
            root.Controls.Add(new Panel { BackColor = MainBackground, Dock = DockStyle.Fill }, 0, 2);
            root.Controls.Add(CreateComposer(), 0, 3);
            root.Controls.Add(CreateStatusBar(), 0, 4);
            Controls.Add(root);
        }

        private Control CreateHeader()
        {
            var header = new Panel { BackColor = MainBackground, Dock = DockStyle.Fill };
            var title = new Label
            {
                BackColor = MainBackground,
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI Semibold", 13.5F, FontStyle.Bold),
                ForeColor = HeaderAccent,
                Height = 30,
                Text = "🤖  MHT AI Assistant",
                TextAlign = ContentAlignment.MiddleLeft
            };

            var badges = new WrapPanelHost
            {
                BackColor = MainBackground,
                Dock = DockStyle.Top,
                Height = 30
            };
            badges.Controls.Add(WrapBadge(_sap2000BadgeLabel));
            badges.Controls.Add(WrapBadge(_etabsBadgeLabel));

            header.Controls.Add(badges);
            header.Controls.Add(_subtitleLabel);
            header.Controls.Add(title);
            return header;
        }

        private static Control WrapBadge(Label label)
        {
            var badge = new BorderedPanel
            {
                BackColor = LightBlueSurface,
                BorderColor = BorderBlue,
                Margin = new Padding(0, 0, 6, 4),
                Padding = new Padding(0),
                Size = new Size(150, 24)
            };
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleCenter;
            badge.Controls.Add(label);
            return badge;
        }

        private Control CreateConversationCard()
        {
            var card = new BorderedPanel
            {
                BackColor = ConversationBackground,
                BorderColor = BorderBlue,
                Dock = DockStyle.Fill,
                Padding = new Padding(0)
            };
            card.Controls.Add(_conversationPanel);
            return card;
        }

        private Control CreateComposer()
        {
            var inputRow = new TableLayoutPanel
            {
                BackColor = MainBackground,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                RowCount = 1
            };
            inputRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inputRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88F));

            var inputCard = new BorderedPanel
            {
                BackColor = Color.White,
                BorderColor = BorderBlue,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 8, 10, 8)
            };
            inputCard.Controls.Add(_inputBox);

            var buttons = new TableLayoutPanel
            {
                BackColor = MainBackground,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 0, 0, 0),
                RowCount = 2
            };
            buttons.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            buttons.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            buttons.Controls.Add(_sendButton, 0, 0);
            buttons.Controls.Add(_clearButton, 0, 1);

            inputRow.Controls.Add(inputCard, 0, 0);
            inputRow.Controls.Add(buttons, 1, 0);
            return inputRow;
        }

        private Control CreateStatusBar()
        {
            var panel = new TableLayoutPanel
            {
                BackColor = MainBackground,
                ColumnCount = 2,
                Dock = DockStyle.Fill
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 126F));

            var right = new Label
            {
                BackColor = MainBackground,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8F),
                ForeColor = MutedInk,
                Text = "Read-only · Ollama local",
                TextAlign = ContentAlignment.MiddleRight
            };

            panel.Controls.Add(_statusLabel, 0, 0);
            panel.Controls.Add(right, 1, 0);
            return panel;
        }

        private void WireEvents()
        {
            _viewModel.Messages.CollectionChanged += Messages_CollectionChanged;
            _viewModel.PropertyChanged += ViewModel_PropertyChanged;
            _viewModel.SendCommand.CanExecuteChanged += Command_CanExecuteChanged;

            _inputBox.TextChanged += InputBox_TextChanged;
            _inputBox.KeyDown += InputBox_KeyDown;
            _sendButton.Click += SendButton_Click;
            _clearButton.Click += ClearButton_Click;
            _conversationPanel.Resize += ConversationPanel_Resize;
        }

        private void Messages_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            RunOnUiThread(UpdateConversation);
        }

        private void ViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            RunOnUiThread(UpdateAll);
        }

        private void Command_CanExecuteChanged(object sender, EventArgs e)
        {
            RunOnUiThread(UpdateButtons);
        }

        private void ConversationPanel_Resize(object sender, EventArgs e)
        {
            UpdateConversation();
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
            UpdateHeader();
            UpdateConversation();
            UpdateButtons();
            _statusLabel.Text = _viewModel.IsBusy ? "Thinking..." : _viewModel.StatusText;
        }

        private void UpdateHeader()
        {
            _subtitleLabel.Text = "Powered by Ollama · Model: " + (_viewModel.CurrentModelName ?? "Not selected") + " · Read-only model access";
            _sap2000BadgeLabel.Text = "SAP2000 Model: " + (_viewModel.Sap2000ConnectionStatus ?? "Attached");
            _etabsBadgeLabel.Text = "ETABS: " + (_viewModel.EtabsConnectionStatus ?? "Attached");
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

        private void UpdateConversation()
        {
            if (_conversationPanel.IsDisposed)
            {
                return;
            }

            _conversationPanel.SuspendLayout();
            _conversationPanel.Controls.Clear();

            if (_viewModel.Messages.Count == 0)
            {
                _conversationPanel.Controls.Add(CreateEmptyState());
            }
            else
            {
                foreach (AiAgentChatMessageViewModel message in _viewModel.Messages)
                {
                    _conversationPanel.Controls.Add(CreateMessageRow(message));
                }
            }

            _conversationPanel.ResumeLayout();
            ScrollConversationToBottom();
        }

        private Control CreateEmptyState()
        {
            return new Label
            {
                AutoSize = false,
                Font = new Font("Segoe UI", 9F, FontStyle.Italic),
                ForeColor = MutedInk,
                Height = 46,
                Margin = new Padding(2, 4, 2, 4),
                Text = "Ask about selected frames, section properties, model units, or current model information.",
                TextAlign = ContentAlignment.MiddleLeft,
                Width = GetConversationWidth()
            };
        }

        private Control CreateMessageRow(AiAgentChatMessageViewModel message)
        {
            int fullWidth = GetConversationWidth();
            int bubbleWidth = Math.Max(180, fullWidth - 56);
            bool isUser = message.IsUser;

            var row = new Panel
            {
                BackColor = ConversationBackground,
                Height = 10,
                Margin = new Padding(0, 0, 0, 8),
                Width = fullWidth
            };

            var bubble = new BorderedPanel
            {
                BackColor = isUser ? HeaderAccent : AssistantBubble,
                BorderColor = isUser ? HeaderAccent : BorderBlue,
                Padding = new Padding(10, 7, 10, 8),
                Width = bubbleWidth
            };

            var label = new Label
            {
                AutoSize = true,
                BackColor = bubble.BackColor,
                Font = new Font("Segoe UI Semibold", 8.25F, FontStyle.Bold),
                ForeColor = isUser ? Color.White : HeaderAccent,
                MaximumSize = new Size(bubbleWidth - 22, 0),
                Text = isUser ? "You" : "MHT AI Assistant"
            };

            var content = new Label
            {
                AutoSize = true,
                BackColor = bubble.BackColor,
                Font = message.IsTemporary
                    ? new Font("Segoe UI", 13F, FontStyle.Bold)
                    : new Font("Segoe UI", 9.25F),
                ForeColor = isUser ? Color.White : Ink,
                MaximumSize = new Size(bubbleWidth - 22, 0),
                Text = message.Content ?? string.Empty,
                Top = label.Bottom + 3
            };

            bubble.Controls.Add(label);
            bubble.Controls.Add(content);
            content.Location = new Point(10, label.Bottom + 5);
            int bubbleHeight = content.Bottom + 9;
            bubble.Height = bubbleHeight;
            bubble.Left = isUser ? fullWidth - bubbleWidth - 6 : 2;
            bubble.Top = 0;
            row.Height = bubbleHeight + 2;
            row.Controls.Add(bubble);
            return row;
        }

        private int GetConversationWidth()
        {
            int width = _conversationPanel.ClientSize.Width - 28;
            return Math.Max(260, width);
        }

        private void ScrollConversationToBottom()
        {
            if (_conversationPanel.Controls.Count == 0)
            {
                return;
            }

            _conversationPanel.ScrollControlIntoView(_conversationPanel.Controls[_conversationPanel.Controls.Count - 1]);
        }

        private void UpdateButtons()
        {
            bool canSend = _viewModel.SendCommand.CanExecute(null);
            _sendButton.Enabled = canSend;
            _sendButton.BackColor = canSend ? HeaderAccent : Color.FromArgb(173, 201, 218);
            _clearButton.Enabled = _viewModel.ClearCommand.CanExecute(null);
            _inputBox.Enabled = !_viewModel.IsBusy;
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
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(BorderColor))
                {
                    Rectangle rect = ClientRectangle;
                    rect.Width -= 1;
                    rect.Height -= 1;
                    e.Graphics.DrawRectangle(pen, rect);
                }
            }
        }

        private sealed class WrapPanelHost : FlowLayoutPanel
        {
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
            }
        }
    }
}
