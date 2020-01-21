namespace Automate_UIP
{
    partial class AutomateUIPRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AutomateUIPRibbon()
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
            this.AutomateUIP = this.Factory.CreateRibbonTab();
            this.CreateGroup = this.Factory.CreateRibbonGroup();
            this.AddNewTradeSheet = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.SendSubUpdates = this.Factory.CreateRibbonSplitButton();
            this.ExportSubReports = this.Factory.CreateRibbonButton();
            this.GenerateFullReport = this.Factory.CreateRibbonButton();
            this.UpdateGroup = this.Factory.CreateRibbonGroup();
            this.OpenQuantityLinkFile = this.Factory.CreateRibbonButton();
            this.RefreshQuantityLinks = this.Factory.CreateRibbonSplitButton();
            this.UpdatePathtoQuantityLinkFile = this.Factory.CreateRibbonButton();
            this.TradeSheetTools = this.Factory.CreateRibbonGroup();
            this.UpdatePlan = this.Factory.CreateRibbonButton();
            this.UpdateWeek = this.Factory.CreateRibbonButton();
            this.SelectTakeoffFiles = this.Factory.CreateRibbonButton();
            this.OtherGroup = this.Factory.CreateRibbonGroup();
            this.Settings = this.Factory.CreateRibbonButton();
            this.AutomateUIP.SuspendLayout();
            this.CreateGroup.SuspendLayout();
            this.UpdateGroup.SuspendLayout();
            this.TradeSheetTools.SuspendLayout();
            this.OtherGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // AutomateUIP
            // 
            this.AutomateUIP.Groups.Add(this.CreateGroup);
            this.AutomateUIP.Groups.Add(this.UpdateGroup);
            this.AutomateUIP.Groups.Add(this.TradeSheetTools);
            this.AutomateUIP.Groups.Add(this.OtherGroup);
            this.AutomateUIP.Label = "Automate UIP";
            this.AutomateUIP.Name = "AutomateUIP";
            // 
            // CreateGroup
            // 
            this.CreateGroup.Items.Add(this.AddNewTradeSheet);
            this.CreateGroup.Items.Add(this.separator1);
            this.CreateGroup.Items.Add(this.SendSubUpdates);
            this.CreateGroup.Items.Add(this.GenerateFullReport);
            this.CreateGroup.Label = "Create / Report";
            this.CreateGroup.Name = "CreateGroup";
            // 
            // AddNewTradeSheet
            // 
            this.AddNewTradeSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddNewTradeSheet.Label = "Add New Trade Sheet";
            this.AddNewTradeSheet.Name = "AddNewTradeSheet";
            this.AddNewTradeSheet.OfficeImageId = "AddAccount";
            this.AddNewTradeSheet.ShowImage = true;
            this.AddNewTradeSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddNewTradeSheet_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // SendSubUpdates
            // 
            this.SendSubUpdates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SendSubUpdates.Items.Add(this.ExportSubReports);
            this.SendSubUpdates.Label = "Export/Send Sub Updates";
            this.SendSubUpdates.Name = "SendSubUpdates";
            this.SendSubUpdates.OfficeImageId = "PublishToPdfOrEdoc";
            this.SendSubUpdates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SendSubUpdates_Click);
            // 
            // ExportSubReports
            // 
            this.ExportSubReports.Label = "Export Sub Reports";
            this.ExportSubReports.Name = "ExportSubReports";
            this.ExportSubReports.OfficeImageId = "FileSaveAsPdfOrXps";
            this.ExportSubReports.ShowImage = true;
            this.ExportSubReports.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportSubReports_Click);
            // 
            // GenerateFullReport
            // 
            this.GenerateFullReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GenerateFullReport.Label = "Generate Full Report";
            this.GenerateFullReport.Name = "GenerateFullReport";
            this.GenerateFullReport.OfficeImageId = "CreateReportInDesignView";
            this.GenerateFullReport.ShowImage = true;
            this.GenerateFullReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateFullReport_Click);
            // 
            // UpdateGroup
            // 
            this.UpdateGroup.Items.Add(this.OpenQuantityLinkFile);
            this.UpdateGroup.Items.Add(this.RefreshQuantityLinks);
            this.UpdateGroup.Label = "Update";
            this.UpdateGroup.Name = "UpdateGroup";
            // 
            // OpenQuantityLinkFile
            // 
            this.OpenQuantityLinkFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OpenQuantityLinkFile.Label = "Open Quantity Link File";
            this.OpenQuantityLinkFile.Name = "OpenQuantityLinkFile";
            this.OpenQuantityLinkFile.OfficeImageId = "HeaderFooterFilePathInsert";
            this.OpenQuantityLinkFile.ShowImage = true;
            this.OpenQuantityLinkFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenQuantityLinkFile_Click);
            // 
            // RefreshQuantityLinks
            // 
            this.RefreshQuantityLinks.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RefreshQuantityLinks.Items.Add(this.UpdatePathtoQuantityLinkFile);
            this.RefreshQuantityLinks.Label = "Refresh Quantity Links";
            this.RefreshQuantityLinks.Name = "RefreshQuantityLinks";
            this.RefreshQuantityLinks.OfficeImageId = "Refresh";
            this.RefreshQuantityLinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RefreshQuantityLinks_Click);
            // 
            // UpdatePathtoQuantityLinkFile
            // 
            this.UpdatePathtoQuantityLinkFile.Label = "Update Path to Quantity Link File";
            this.UpdatePathtoQuantityLinkFile.Name = "UpdatePathtoQuantityLinkFile";
            this.UpdatePathtoQuantityLinkFile.OfficeImageId = "FileFind";
            this.UpdatePathtoQuantityLinkFile.ShowImage = true;
            this.UpdatePathtoQuantityLinkFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdatePathtoQuantityLinkFile_Click);
            // 
            // TradeSheetTools
            // 
            this.TradeSheetTools.Items.Add(this.UpdatePlan);
            this.TradeSheetTools.Items.Add(this.UpdateWeek);
            this.TradeSheetTools.Items.Add(this.SelectTakeoffFiles);
            this.TradeSheetTools.Label = "Trade Sheet Tools";
            this.TradeSheetTools.Name = "TradeSheetTools";
            // 
            // UpdatePlan
            // 
            this.UpdatePlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UpdatePlan.Label = "Update Plan";
            this.UpdatePlan.Name = "UpdatePlan";
            this.UpdatePlan.OfficeImageId = "GroupChartData";
            this.UpdatePlan.ShowImage = true;
            this.UpdatePlan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdatePlan_Click);
            // 
            // UpdateWeek
            // 
            this.UpdateWeek.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UpdateWeek.Label = "Update Week";
            this.UpdateWeek.Name = "UpdateWeek";
            this.UpdateWeek.OfficeImageId = "ChartRefresh";
            this.UpdateWeek.ShowImage = true;
            this.UpdateWeek.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateWeek_Click);
            // 
            // SelectTakeoffFiles
            // 
            this.SelectTakeoffFiles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SelectTakeoffFiles.Label = "Select Takeoff Files";
            this.SelectTakeoffFiles.Name = "SelectTakeoffFiles";
            this.SelectTakeoffFiles.OfficeImageId = "FileFind";
            this.SelectTakeoffFiles.ShowImage = true;
            this.SelectTakeoffFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectTakeoffFiles_Click);
            // 
            // OtherGroup
            // 
            this.OtherGroup.Items.Add(this.Settings);
            this.OtherGroup.Label = "Other";
            this.OtherGroup.Name = "OtherGroup";
            // 
            // Settings
            // 
            this.Settings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Settings.Label = "Settings";
            this.Settings.Name = "Settings";
            this.Settings.OfficeImageId = "AdministrationHome";
            this.Settings.ShowImage = true;
            this.Settings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Settings_Click);
            // 
            // AutomateUIPRibbon
            // 
            this.Name = "AutomateUIPRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.AutomateUIP);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.AutomateUIP.ResumeLayout(false);
            this.AutomateUIP.PerformLayout();
            this.CreateGroup.ResumeLayout(false);
            this.CreateGroup.PerformLayout();
            this.UpdateGroup.ResumeLayout(false);
            this.UpdateGroup.PerformLayout();
            this.TradeSheetTools.ResumeLayout(false);
            this.TradeSheetTools.PerformLayout();
            this.OtherGroup.ResumeLayout(false);
            this.OtherGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AutomateUIP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CreateGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddNewTradeSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup UpdateGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenQuantityLinkFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateFullReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton RefreshQuantityLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdatePathtoQuantityLinkFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TradeSheetTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdatePlan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateWeek;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SendSubUpdates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportSubReports;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SelectTakeoffFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup OtherGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Settings;
    }

    partial class ThisRibbonCollection
    {
        internal AutomateUIPRibbon Ribbon1
        {
            get { return this.GetRibbon<AutomateUIPRibbon>(); }
        }
    }
}
