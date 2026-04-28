namespace ExcelCSIToolBoxAddIn
{
    partial class ExcelCSIToolBoxAddInRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelCSIToolBoxAddInRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabExcelCSIToolBox = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonEtabs = this.Factory.CreateRibbonButton();
            this.buttonSap2000 = this.Factory.CreateRibbonButton();
            this.tabExcelCSIToolBox.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabExcelCSIToolBox
            // 
            this.tabExcelCSIToolBox.Groups.Add(this.group1);
            this.tabExcelCSIToolBox.Label = "ExcelCSIToolBox";
            this.tabExcelCSIToolBox.Name = "tabExcelCSIToolBox";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonEtabs);
            this.group1.Items.Add(this.buttonSap2000);
            this.group1.Label = "CSI Toolbox";
            this.group1.Name = "group1";
            // 
            // buttonEtabs
            // 
            this.buttonEtabs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonEtabs.Label = "ETABS Toolbox";
            this.buttonEtabs.Name = "buttonEtabs";
            this.buttonEtabs.ShowImage = true;
            this.buttonEtabs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEtabs_Click);
            // 
            // buttonSap2000
            // 
            this.buttonSap2000.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSap2000.Label = "SAP2000 Toolbox";
            this.buttonSap2000.Name = "buttonSap2000";
            this.buttonSap2000.ShowImage = true;
            this.buttonSap2000.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSap2000_Click);
            // 
            // ExcelCSIToolBoxAddInRibbon
            // 
            this.Name = "ExcelCSIToolBoxAddInRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabExcelCSIToolBox);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabExcelCSIToolBox.ResumeLayout(false);
            this.tabExcelCSIToolBox.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcelCSIToolBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEtabs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSap2000;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelCSIToolBoxAddInRibbon ExcelCSIToolBoxAddIn
        {
            get { return this.GetRibbon<ExcelCSIToolBoxAddInRibbon>(); }
        }
    }
}
