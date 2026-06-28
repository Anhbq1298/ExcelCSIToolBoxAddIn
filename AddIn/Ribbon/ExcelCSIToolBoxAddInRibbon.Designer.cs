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
            this.groupEtabsPostprocessing = this.Factory.CreateRibbonGroup();
            this.buttonGetBaseReactions = this.Factory.CreateRibbonButton();
            this.buttonModalMassParticipationRatios = this.Factory.CreateRibbonButton();
            this.buttonStoryForces = this.Factory.CreateRibbonButton();
            this.buttonStoryDrifts = this.Factory.CreateRibbonButton();
            this.buttonStoryMaxOverAverageDisplacements = this.Factory.CreateRibbonButton();
            this.buttonStoryMaxOverAverageDrifts = this.Factory.CreateRibbonButton();
            this.buttonMassSummaryByStory = this.Factory.CreateRibbonButton();
            this.groupAiAssistant = this.Factory.CreateRibbonGroup();
            this.buttonAiAgent = this.Factory.CreateRibbonButton();
            this.tabExcelCSIToolBox.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupEtabsPostprocessing.SuspendLayout();
            this.groupAiAssistant.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabExcelCSIToolBox
            // 
            this.tabExcelCSIToolBox.Groups.Add(this.group1);
            this.tabExcelCSIToolBox.Groups.Add(this.groupEtabsPostprocessing);
            this.tabExcelCSIToolBox.Groups.Add(this.groupAiAssistant);
            this.tabExcelCSIToolBox.Label = "ExcelCSIToolBox";
            this.tabExcelCSIToolBox.Name = "tabExcelCSIToolBox";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonEtabs);
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
            // groupEtabsPostprocessing
            // 
            this.groupEtabsPostprocessing.Items.Add(this.buttonGetBaseReactions);
            this.groupEtabsPostprocessing.Items.Add(this.buttonModalMassParticipationRatios);
            this.groupEtabsPostprocessing.Items.Add(this.buttonStoryForces);
            this.groupEtabsPostprocessing.Items.Add(this.buttonStoryDrifts);
            this.groupEtabsPostprocessing.Items.Add(this.buttonStoryMaxOverAverageDisplacements);
            this.groupEtabsPostprocessing.Items.Add(this.buttonStoryMaxOverAverageDrifts);
            this.groupEtabsPostprocessing.Items.Add(this.buttonMassSummaryByStory);
            this.groupEtabsPostprocessing.Label = "ETABS postprocessing toolbox";
            this.groupEtabsPostprocessing.Name = "groupEtabsPostprocessing";
            // 
            // buttonGetBaseReactions
            // 
            this.buttonGetBaseReactions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonGetBaseReactions.Label = "Get Base Reactions";
            this.buttonGetBaseReactions.Name = "buttonGetBaseReactions";
            this.buttonGetBaseReactions.OfficeImageId = "TableExport";
            this.buttonGetBaseReactions.ScreenTip = "Get Base Reactions";
            this.buttonGetBaseReactions.ShowImage = true;
            this.buttonGetBaseReactions.SuperTip = "Extract ETABS Base Reactions for selected load cases and combinations to Excel.";
            this.buttonGetBaseReactions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetBaseReactions_Click);
            // 
            // buttonModalMassParticipationRatios
            // 
            this.buttonModalMassParticipationRatios.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonModalMassParticipationRatios.Label = "Modal Mass Participation Ratios";
            this.buttonModalMassParticipationRatios.Name = "buttonModalMassParticipationRatios";
            this.buttonModalMassParticipationRatios.OfficeImageId = "TableOfContentsInsert";
            this.buttonModalMassParticipationRatios.ScreenTip = "Modal Mass Participation Ratios";
            this.buttonModalMassParticipationRatios.ShowImage = true;
            this.buttonModalMassParticipationRatios.SuperTip = "Extract ETABS Modal Mass Participation Ratios for selected modal load cases to Excel.";
            this.buttonModalMassParticipationRatios.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonModalMassParticipationRatios_Click);
            // 
            // buttonStoryForces
            // 
            this.buttonStoryForces.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStoryForces.Label = "Story Forces";
            this.buttonStoryForces.Name = "buttonStoryForces";
            this.buttonStoryForces.OfficeImageId = "TableExport";
            this.buttonStoryForces.ScreenTip = "Story Forces";
            this.buttonStoryForces.ShowImage = true;
            this.buttonStoryForces.SuperTip = "Extract ETABS Story Forces for selected load cases and combinations to Excel.";
            this.buttonStoryForces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStoryForces_Click);
            // 
            // buttonStoryDrifts
            // 
            this.buttonStoryDrifts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStoryDrifts.Label = "Story Drifts";
            this.buttonStoryDrifts.Name = "buttonStoryDrifts";
            this.buttonStoryDrifts.OfficeImageId = "TableExport";
            this.buttonStoryDrifts.ScreenTip = "Story Drifts";
            this.buttonStoryDrifts.ShowImage = true;
            this.buttonStoryDrifts.SuperTip = "Extract ETABS Story Drifts for selected load cases and combinations to Excel.";
            this.buttonStoryDrifts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStoryDrifts_Click);
            // 
            // buttonStoryMaxOverAverageDisplacements
            // 
            this.buttonStoryMaxOverAverageDisplacements.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStoryMaxOverAverageDisplacements.Label = "Story Max/Avg Displacements";
            this.buttonStoryMaxOverAverageDisplacements.Name = "buttonStoryMaxOverAverageDisplacements";
            this.buttonStoryMaxOverAverageDisplacements.OfficeImageId = "TableExport";
            this.buttonStoryMaxOverAverageDisplacements.ScreenTip = "Story Max Over Avg Displacements";
            this.buttonStoryMaxOverAverageDisplacements.ShowImage = true;
            this.buttonStoryMaxOverAverageDisplacements.SuperTip = "Extract ETABS Story Max Over Avg Displacements for selected load cases and combinations to Excel.";
            this.buttonStoryMaxOverAverageDisplacements.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStoryMaxOverAverageDisplacements_Click);
            // 
            // buttonStoryMaxOverAverageDrifts
            // 
            this.buttonStoryMaxOverAverageDrifts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonStoryMaxOverAverageDrifts.Label = "Story Max/Avg Drifts";
            this.buttonStoryMaxOverAverageDrifts.Name = "buttonStoryMaxOverAverageDrifts";
            this.buttonStoryMaxOverAverageDrifts.OfficeImageId = "TableExport";
            this.buttonStoryMaxOverAverageDrifts.ScreenTip = "Story Max Over Avg Drifts";
            this.buttonStoryMaxOverAverageDrifts.ShowImage = true;
            this.buttonStoryMaxOverAverageDrifts.SuperTip = "Extract ETABS Story Max Over Avg Drifts for selected load cases and combinations to Excel.";
            this.buttonStoryMaxOverAverageDrifts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStoryMaxOverAverageDrifts_Click);
            // 
            // buttonMassSummaryByStory
            // 
            this.buttonMassSummaryByStory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonMassSummaryByStory.Label = "Mass Summary by Story";
            this.buttonMassSummaryByStory.Name = "buttonMassSummaryByStory";
            this.buttonMassSummaryByStory.OfficeImageId = "TableExport";
            this.buttonMassSummaryByStory.ScreenTip = "Mass Summary by Story";
            this.buttonMassSummaryByStory.ShowImage = true;
            this.buttonMassSummaryByStory.SuperTip = "Extract ETABS Mass Summary by Story to Excel.";
            this.buttonMassSummaryByStory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMassSummaryByStory_Click);
            // 
            // groupAiAssistant
            // 
            this.groupAiAssistant.Items.Add(this.buttonAiAgent);
            this.groupAiAssistant.Label = "AI Assistant";
            this.groupAiAssistant.Name = "groupAiAssistant";
            // 
            // buttonAiAgent
            // 
            this.buttonAiAgent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAiAgent.Label = "MHT AI Assistant";
            this.buttonAiAgent.Name = "buttonAiAgent";
            this.buttonAiAgent.OfficeImageId = "HappyFace";
            this.buttonAiAgent.ScreenTip = "Open MHT AI Assistant";
            this.buttonAiAgent.ShowImage = true;
            this.buttonAiAgent.SuperTip = "Open the local Ollama-powered AI assistant in an Excel custom task pane.";
            this.buttonAiAgent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAiAgent_Click);
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
            this.groupEtabsPostprocessing.ResumeLayout(false);
            this.groupEtabsPostprocessing.PerformLayout();
            this.groupAiAssistant.ResumeLayout(false);
            this.groupAiAssistant.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcelCSIToolBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEtabsPostprocessing;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAiAssistant;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEtabs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetBaseReactions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonModalMassParticipationRatios;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStoryForces;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStoryDrifts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStoryMaxOverAverageDisplacements;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStoryMaxOverAverageDrifts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMassSummaryByStory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAiAgent;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelCSIToolBoxAddInRibbon ExcelCSIToolBoxAddIn
        {
            get { return this.GetRibbon<ExcelCSIToolBoxAddInRibbon>(); }
        }
    }
}

