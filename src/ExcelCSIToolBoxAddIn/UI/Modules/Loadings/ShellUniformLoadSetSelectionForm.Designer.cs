namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    partial class ShellUniformLoadSetSelectionForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblFilter;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.Label lblStory;
        private System.Windows.Forms.CheckedListBox chkStories;
        private System.Windows.Forms.CheckedListBox chkLoadSets;
        private System.Windows.Forms.Label lblEmptyState;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Button btnClearSelection;
        private System.Windows.Forms.Button btnSelectAllStories;
        private System.Windows.Forms.Button btnClearStories;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSelectShells;
        private System.Windows.Forms.Label lblSelectionStatus;
        private System.Windows.Forms.Label lblStorySelectionStatus;
        private System.Windows.Forms.ProgressBar progressSelection;
        private System.Windows.Forms.Label lblStatus;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblFilter = new System.Windows.Forms.Label();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.lblStory = new System.Windows.Forms.Label();
            this.chkStories = new System.Windows.Forms.CheckedListBox();
            this.chkLoadSets = new System.Windows.Forms.CheckedListBox();
            this.lblEmptyState = new System.Windows.Forms.Label();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnClearSelection = new System.Windows.Forms.Button();
            this.btnSelectAllStories = new System.Windows.Forms.Button();
            this.btnClearStories = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSelectShells = new System.Windows.Forms.Button();
            this.lblSelectionStatus = new System.Windows.Forms.Label();
            this.lblStorySelectionStatus = new System.Windows.Forms.Label();
            this.progressSelection = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            this.lblFilter.AutoSize = true;
            this.lblFilter.Location = new System.Drawing.Point(16, 18);
            this.lblFilter.Name = "lblFilter";
            this.lblFilter.Size = new System.Drawing.Size(99, 13);
            this.lblFilter.TabIndex = 0;
            this.lblFilter.Text = "Filter Load Set(s):";
            this.txtFilter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)));
            this.txtFilter.Location = new System.Drawing.Point(19, 40);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(370, 20);
            this.txtFilter.TabIndex = 1;
            this.txtFilter.TextChanged += new System.EventHandler(this.txtFilter_TextChanged);
            this.lblStory.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStory.AutoSize = true;
            this.lblStory.Location = new System.Drawing.Point(410, 18);
            this.lblStory.Name = "lblStory";
            this.lblStory.Size = new System.Drawing.Size(48, 13);
            this.lblStory.TabIndex = 2;
            this.lblStory.Text = "Story(s):";
            this.chkStories.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chkStories.CheckOnClick = true;
            this.chkStories.FormattingEnabled = true;
            this.chkStories.IntegralHeight = false;
            this.chkStories.Location = new System.Drawing.Point(413, 40);
            this.chkStories.Name = "chkStories";
            this.chkStories.Size = new System.Drawing.Size(328, 320);
            this.chkStories.TabIndex = 3;
            this.chkStories.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkStories_ItemCheck);
            this.chkLoadSets.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)));
            this.chkLoadSets.CheckOnClick = true;
            this.chkLoadSets.FormattingEnabled = true;
            this.chkLoadSets.IntegralHeight = false;
            this.chkLoadSets.Location = new System.Drawing.Point(19, 76);
            this.chkLoadSets.Name = "chkLoadSets";
            this.chkLoadSets.Size = new System.Drawing.Size(370, 284);
            this.chkLoadSets.TabIndex = 4;
            this.chkLoadSets.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkLoadSets_ItemCheck);
            this.lblEmptyState.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)));
            this.lblEmptyState.AutoEllipsis = true;
            this.lblEmptyState.BackColor = System.Drawing.SystemColors.Window;
            this.lblEmptyState.ForeColor = System.Drawing.SystemColors.GrayText;
            this.lblEmptyState.Location = new System.Drawing.Point(34, 91);
            this.lblEmptyState.Name = "lblEmptyState";
            this.lblEmptyState.Size = new System.Drawing.Size(340, 40);
            this.lblEmptyState.TabIndex = 5;
            this.lblEmptyState.Text = "No Shell Uniform Load Sets exist in the connected ETABS model.";
            this.lblEmptyState.Visible = false;
            this.btnSelectAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSelectAll.Location = new System.Drawing.Point(19, 375);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(88, 28);
            this.btnSelectAll.TabIndex = 6;
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            this.btnClearSelection.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClearSelection.Location = new System.Drawing.Point(113, 375);
            this.btnClearSelection.Name = "btnClearSelection";
            this.btnClearSelection.Size = new System.Drawing.Size(112, 28);
            this.btnClearSelection.TabIndex = 7;
            this.btnClearSelection.Text = "Clear Selection";
            this.btnClearSelection.UseVisualStyleBackColor = true;
            this.btnClearSelection.Click += new System.EventHandler(this.btnClearSelection_Click);
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnRefresh.Location = new System.Drawing.Point(231, 375);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(88, 28);
            this.btnRefresh.TabIndex = 8;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            this.btnSelectAllStories.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectAllStories.Location = new System.Drawing.Point(413, 375);
            this.btnSelectAllStories.Name = "btnSelectAllStories";
            this.btnSelectAllStories.Size = new System.Drawing.Size(88, 28);
            this.btnSelectAllStories.TabIndex = 9;
            this.btnSelectAllStories.Text = "Select All";
            this.btnSelectAllStories.UseVisualStyleBackColor = true;
            this.btnSelectAllStories.Click += new System.EventHandler(this.btnSelectAllStories_Click);
            this.btnClearStories.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClearStories.Location = new System.Drawing.Point(507, 375);
            this.btnClearStories.Name = "btnClearStories";
            this.btnClearStories.Size = new System.Drawing.Size(88, 28);
            this.btnClearStories.TabIndex = 10;
            this.btnClearStories.Text = "Clear";
            this.btnClearStories.UseVisualStyleBackColor = true;
            this.btnClearStories.Click += new System.EventHandler(this.btnClearStories_Click);
            this.lblSelectionStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblSelectionStatus.AutoEllipsis = true;
            this.lblSelectionStatus.Location = new System.Drawing.Point(16, 412);
            this.lblSelectionStatus.Name = "lblSelectionStatus";
            this.lblSelectionStatus.Size = new System.Drawing.Size(373, 18);
            this.lblSelectionStatus.TabIndex = 11;
            this.lblSelectionStatus.Text = "0 load set(s) selected";
            this.lblSelectionStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblStorySelectionStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStorySelectionStatus.AutoEllipsis = true;
            this.lblStorySelectionStatus.Location = new System.Drawing.Point(410, 412);
            this.lblStorySelectionStatus.Name = "lblStorySelectionStatus";
            this.lblStorySelectionStatus.Size = new System.Drawing.Size(331, 18);
            this.lblStorySelectionStatus.TabIndex = 12;
            this.lblStorySelectionStatus.Text = "0 story(s) selected";
            this.lblStorySelectionStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.progressSelection.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressSelection.Location = new System.Drawing.Point(19, 434);
            this.progressSelection.Name = "progressSelection";
            this.progressSelection.Size = new System.Drawing.Size(523, 12);
            this.progressSelection.TabIndex = 13;
            this.progressSelection.Visible = false;
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.Location = new System.Drawing.Point(16, 456);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(526, 23);
            this.lblStatus.TabIndex = 14;
            this.lblStatus.Text = "Ready.";
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(563, 451);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(86, 30);
            this.btnCancel.TabIndex = 15;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnSelectShells.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectShells.Enabled = false;
            this.btnSelectShells.Location = new System.Drawing.Point(655, 451);
            this.btnSelectShells.Name = "btnSelectShells";
            this.btnSelectShells.Size = new System.Drawing.Size(90, 30);
            this.btnSelectShells.TabIndex = 16;
            this.btnSelectShells.Text = "Select Shells";
            this.btnSelectShells.UseVisualStyleBackColor = true;
            this.btnSelectShells.Click += new System.EventHandler(this.btnSelectShells_Click);
            this.AcceptButton = this.btnSelectShells;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(760, 501);
            this.Controls.Add(this.btnSelectShells);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressSelection);
            this.Controls.Add(this.lblStorySelectionStatus);
            this.Controls.Add(this.lblSelectionStatus);
            this.Controls.Add(this.btnClearStories);
            this.Controls.Add(this.btnSelectAllStories);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.btnClearSelection);
            this.Controls.Add(this.btnSelectAll);
            this.Controls.Add(this.lblEmptyState);
            this.Controls.Add(this.chkLoadSets);
            this.Controls.Add(this.chkStories);
            this.Controls.Add(this.lblStory);
            this.Controls.Add(this.txtFilter);
            this.Controls.Add(this.lblFilter);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(720, 430);
            this.Name = "ShellUniformLoadSetSelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Select Shells by Load Set";
            this.Load += new System.EventHandler(this.ShellUniformLoadSetSelectionForm_Load);
            this.Shown += new System.EventHandler(this.ShellUniformLoadSetSelectionForm_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
