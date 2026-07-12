namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    partial class ShellUniformLoadSetForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblModelName;
        private System.Windows.Forms.Label lblPresentUnits;
        private System.Windows.Forms.Label lblLoadValues;
        private System.Windows.Forms.Label lblSelectedUnits;
        private System.Windows.Forms.ComboBox cboSelectedUnits;
        private System.Windows.Forms.CheckBox chkApplySuggestedName;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnImportExcelRange;
        private System.Windows.Forms.Button btnRefreshDefinitions;
        private System.Windows.Forms.Button btnExportDefinitions;
        private System.Windows.Forms.Button btnAddRow;
        private System.Windows.Forms.Button btnDeleteRow;
        private System.Windows.Forms.Button btnAddLoadPattern;
        private System.Windows.Forms.Button btnRemoveLoadPattern;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.DataGridView dgvLoadSets;

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
            this.lblModelName = new System.Windows.Forms.Label();
            this.lblPresentUnits = new System.Windows.Forms.Label();
            this.lblLoadValues = new System.Windows.Forms.Label();
            this.lblSelectedUnits = new System.Windows.Forms.Label();
            this.cboSelectedUnits = new System.Windows.Forms.ComboBox();
            this.chkApplySuggestedName = new System.Windows.Forms.CheckBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnImportExcelRange = new System.Windows.Forms.Button();
            this.btnRefreshDefinitions = new System.Windows.Forms.Button();
            this.btnExportDefinitions = new System.Windows.Forms.Button();
            this.btnAddRow = new System.Windows.Forms.Button();
            this.btnDeleteRow = new System.Windows.Forms.Button();
            this.btnAddLoadPattern = new System.Windows.Forms.Button();
            this.btnRemoveLoadPattern = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.dgvLoadSets = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLoadSets)).BeginInit();
            this.SuspendLayout();
            this.lblModelName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblModelName.AutoEllipsis = true;
            this.lblModelName.Location = new System.Drawing.Point(16, 14);
            this.lblModelName.Name = "lblModelName";
            this.lblModelName.Size = new System.Drawing.Size(952, 20);
            this.lblModelName.TabIndex = 0;
            this.lblModelName.Text = "ETABS Model: -";
            this.lblPresentUnits.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPresentUnits.AutoEllipsis = true;
            this.lblPresentUnits.Location = new System.Drawing.Point(16, 38);
            this.lblPresentUnits.Name = "lblPresentUnits";
            this.lblPresentUnits.Size = new System.Drawing.Size(952, 20);
            this.lblPresentUnits.TabIndex = 1;
            this.lblPresentUnits.Text = "Present Units: -";
            this.lblLoadValues.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblLoadValues.AutoEllipsis = true;
            this.lblLoadValues.Location = new System.Drawing.Point(16, 62);
            this.lblLoadValues.Name = "lblLoadValues";
            this.lblLoadValues.Size = new System.Drawing.Size(952, 20);
            this.lblLoadValues.TabIndex = 2;
            this.lblLoadValues.Text = "Load Values: Selected ETABS table units";
            this.lblSelectedUnits.AutoSize = true;
            this.lblSelectedUnits.Location = new System.Drawing.Point(16, 94);
            this.lblSelectedUnits.Name = "lblSelectedUnits";
            this.lblSelectedUnits.Size = new System.Drawing.Size(63, 13);
            this.lblSelectedUnits.TabIndex = 3;
            this.lblSelectedUnits.Text = "Apply Units:";
            this.cboSelectedUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSelectedUnits.FormattingEnabled = true;
            this.cboSelectedUnits.Location = new System.Drawing.Point(85, 91);
            this.cboSelectedUnits.Name = "cboSelectedUnits";
            this.cboSelectedUnits.Size = new System.Drawing.Size(110, 21);
            this.cboSelectedUnits.TabIndex = 4;
            this.cboSelectedUnits.SelectedIndexChanged += new System.EventHandler(this.cboSelectedUnits_SelectedIndexChanged);
            this.btnImportExcelRange.Location = new System.Drawing.Point(219, 86);
            this.btnImportExcelRange.Name = "btnImportExcelRange";
            this.btnImportExcelRange.Size = new System.Drawing.Size(185, 30);
            this.btnImportExcelRange.TabIndex = 5;
            this.btnImportExcelRange.Text = "Import Selected Excel Range";
            this.btnImportExcelRange.UseVisualStyleBackColor = true;
            this.btnImportExcelRange.Click += new System.EventHandler(this.btnImportExcelRange_Click);
            this.chkApplySuggestedName.AutoSize = true;
            this.chkApplySuggestedName.Location = new System.Drawing.Point(420, 93);
            this.chkApplySuggestedName.Name = "chkApplySuggestedName";
            this.chkApplySuggestedName.Size = new System.Drawing.Size(139, 17);
            this.chkApplySuggestedName.TabIndex = 6;
            this.chkApplySuggestedName.Text = "Apply Suggested Name";
            this.chkApplySuggestedName.UseVisualStyleBackColor = true;
            this.chkApplySuggestedName.CheckedChanged += new System.EventHandler(this.chkApplySuggestedName_CheckedChanged);
            this.btnRefreshDefinitions.Location = new System.Drawing.Point(219, 121);
            this.btnRefreshDefinitions.Name = "btnRefreshDefinitions";
            this.btnRefreshDefinitions.Size = new System.Drawing.Size(180, 30);
            this.btnRefreshDefinitions.TabIndex = 7;
            this.btnRefreshDefinitions.Text = "Refresh Current Definitions";
            this.btnRefreshDefinitions.UseVisualStyleBackColor = true;
            this.btnRefreshDefinitions.Click += new System.EventHandler(this.btnRefreshDefinitions_Click);
            this.btnExportDefinitions.Location = new System.Drawing.Point(405, 121);
            this.btnExportDefinitions.Name = "btnExportDefinitions";
            this.btnExportDefinitions.Size = new System.Drawing.Size(320, 30);
            this.btnExportDefinitions.TabIndex = 8;
            this.btnExportDefinitions.Text = "Export Current Shell Uniform Load Set Definitions";
            this.btnExportDefinitions.UseVisualStyleBackColor = true;
            this.btnExportDefinitions.Click += new System.EventHandler(this.btnExportDefinitions_Click);
            this.dgvLoadSets.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvLoadSets.AllowUserToAddRows = false;
            this.dgvLoadSets.AllowUserToOrderColumns = false;
            this.dgvLoadSets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvLoadSets.Location = new System.Drawing.Point(19, 167);
            this.dgvLoadSets.Name = "dgvLoadSets";
            this.dgvLoadSets.RowHeadersWidth = 46;
            this.dgvLoadSets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvLoadSets.Size = new System.Drawing.Size(949, 304);
            this.dgvLoadSets.TabIndex = 9;
            this.dgvLoadSets.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgvLoadSets_CellBeginEdit);
            this.dgvLoadSets.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvLoadSets_CellEndEdit);
            this.btnAddRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAddRow.Location = new System.Drawing.Point(19, 484);
            this.btnAddRow.Name = "btnAddRow";
            this.btnAddRow.Size = new System.Drawing.Size(85, 28);
            this.btnAddRow.TabIndex = 10;
            this.btnAddRow.Text = "+ Add Row";
            this.btnAddRow.UseVisualStyleBackColor = true;
            this.btnAddRow.Click += new System.EventHandler(this.btnAddRow_Click);
            this.btnDeleteRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDeleteRow.Location = new System.Drawing.Point(110, 484);
            this.btnDeleteRow.Name = "btnDeleteRow";
            this.btnDeleteRow.Size = new System.Drawing.Size(85, 28);
            this.btnDeleteRow.TabIndex = 11;
            this.btnDeleteRow.Text = "Delete Row";
            this.btnDeleteRow.UseVisualStyleBackColor = true;
            this.btnDeleteRow.Click += new System.EventHandler(this.btnDeleteRow_Click);
            this.btnAddLoadPattern.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAddLoadPattern.Location = new System.Drawing.Point(201, 484);
            this.btnAddLoadPattern.Name = "btnAddLoadPattern";
            this.btnAddLoadPattern.Size = new System.Drawing.Size(130, 28);
            this.btnAddLoadPattern.TabIndex = 12;
            this.btnAddLoadPattern.Text = "+ Add Load Pattern";
            this.btnAddLoadPattern.UseVisualStyleBackColor = true;
            this.btnAddLoadPattern.Click += new System.EventHandler(this.btnAddLoadPattern_Click);
            this.btnRemoveLoadPattern.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnRemoveLoadPattern.Location = new System.Drawing.Point(337, 484);
            this.btnRemoveLoadPattern.Name = "btnRemoveLoadPattern";
            this.btnRemoveLoadPattern.Size = new System.Drawing.Size(140, 28);
            this.btnRemoveLoadPattern.TabIndex = 13;
            this.btnRemoveLoadPattern.Text = "Remove Load Pattern";
            this.btnRemoveLoadPattern.UseVisualStyleBackColor = true;
            this.btnRemoveLoadPattern.Click += new System.EventHandler(this.btnRemoveLoadPattern_Click);
            this.btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClear.Location = new System.Drawing.Point(483, 484);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 28);
            this.btnClear.TabIndex = 14;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.Location = new System.Drawing.Point(16, 529);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(614, 23);
            this.lblStatus.TabIndex = 15;
            this.lblStatus.Text = "Status: Ready.";
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(783, 524);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 30);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnApply.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnApply.Location = new System.Drawing.Point(877, 524);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(91, 30);
            this.btnApply.TabIndex = 17;
            this.btnApply.Text = "Apply to ETABS";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            this.AcceptButton = this.btnApply;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(984, 571);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnRemoveLoadPattern);
            this.Controls.Add(this.btnAddLoadPattern);
            this.Controls.Add(this.btnDeleteRow);
            this.Controls.Add(this.btnAddRow);
            this.Controls.Add(this.dgvLoadSets);
            this.Controls.Add(this.btnExportDefinitions);
            this.Controls.Add(this.btnRefreshDefinitions);
            this.Controls.Add(this.btnImportExcelRange);
            this.Controls.Add(this.chkApplySuggestedName);
            this.Controls.Add(this.cboSelectedUnits);
            this.Controls.Add(this.lblSelectedUnits);
            this.Controls.Add(this.lblLoadValues);
            this.Controls.Add(this.lblPresentUnits);
            this.Controls.Add(this.lblModelName);
            this.MinimumSize = new System.Drawing.Size(800, 480);
            this.Name = "ShellUniformLoadSetForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Shell Uniform Load Set Manager";
            this.Load += new System.EventHandler(this.ShellUniformLoadSetForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLoadSets)).EndInit();
            this.ResumeLayout(false);
        }
    }
}
