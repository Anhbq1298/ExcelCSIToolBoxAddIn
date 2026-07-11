namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    partial class LoadPatternPickerForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblLoadPattern;
        private System.Windows.Forms.ComboBox cboLoadPatterns;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAdd;

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
            this.lblLoadPattern = new System.Windows.Forms.Label();
            this.cboLoadPatterns = new System.Windows.Forms.ComboBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.SuspendLayout();
            this.lblLoadPattern.AutoSize = true;
            this.lblLoadPattern.Location = new System.Drawing.Point(16, 18);
            this.lblLoadPattern.Name = "lblLoadPattern";
            this.lblLoadPattern.Size = new System.Drawing.Size(75, 13);
            this.lblLoadPattern.TabIndex = 0;
            this.lblLoadPattern.Text = "Load Pattern:";
            this.cboLoadPatterns.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLoadPatterns.FormattingEnabled = true;
            this.cboLoadPatterns.Location = new System.Drawing.Point(19, 42);
            this.cboLoadPatterns.Name = "cboLoadPatterns";
            this.cboLoadPatterns.Size = new System.Drawing.Size(290, 21);
            this.cboLoadPatterns.TabIndex = 1;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(153, 87);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 26);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnAdd.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnAdd.Location = new System.Drawing.Point(234, 87);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 26);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.AcceptButton = this.btnAdd;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(331, 132);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cboLoadPatterns);
            this.Controls.Add(this.lblLoadPattern);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LoadPatternPickerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Add Load Pattern";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
