namespace SpamGrabber
{
    partial class SpamGrabber_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SpamGrabber_Ribbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Report = this.Factory.CreateRibbonGroup();
            this.ButtonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.btnReportDefaultSpam = this.Factory.CreateRibbonButton();
            this.btnReportDefaultHam = this.Factory.CreateRibbonButton();
            this.boxReportTo = this.Factory.CreateRibbonBox();
            this.ddlReportTo = this.Factory.CreateRibbonDropDown();
            this.btnReportCustom = this.Factory.CreateRibbonButton();
            this.grSource = this.Factory.CreateRibbonGroup();
            this.btnSafeView = this.Factory.CreateRibbonButton();
            this.btnCopyToClipboard = this.Factory.CreateRibbonButton();
            this.gpSettings = this.Factory.CreateRibbonGroup();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Report.SuspendLayout();
            this.ButtonGroup1.SuspendLayout();
            this.boxReportTo.SuspendLayout();
            this.grSource.SuspendLayout();
            this.gpSettings.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Report);
            this.tab1.Groups.Add(this.grSource);
            this.tab1.Groups.Add(this.gpSettings);
            this.tab1.Label = "SpamGrabber";
            this.tab1.Name = "tab1";
            // 
            // Report
            // 
            this.Report.Items.Add(this.ButtonGroup1);
            this.Report.Items.Add(this.boxReportTo);
            this.Report.Label = "Report";
            this.Report.Name = "Report";
            // 
            // ButtonGroup1
            // 
            this.ButtonGroup1.Items.Add(this.btnReportDefaultSpam);
            this.ButtonGroup1.Items.Add(this.btnReportDefaultHam);
            this.ButtonGroup1.Name = "ButtonGroup1";
            // 
            // btnReportDefaultSpam
            // 
            this.btnReportDefaultSpam.Image = global::SpamGrabber.Properties.Resources.spamgrab_red;
            this.btnReportDefaultSpam.Label = "Report Spam";
            this.btnReportDefaultSpam.Name = "btnReportDefaultSpam";
            this.btnReportDefaultSpam.ShowImage = true;
            this.btnReportDefaultSpam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReportDefaultSpam_Click);
            // 
            // btnReportDefaultHam
            // 
            this.btnReportDefaultHam.Image = global::SpamGrabber.Properties.Resources.spamgrab_green;
            this.btnReportDefaultHam.Label = "Report Ham";
            this.btnReportDefaultHam.Name = "btnReportDefaultHam";
            this.btnReportDefaultHam.ShowImage = true;
            // 
            // boxReportTo
            // 
            this.boxReportTo.Items.Add(this.ddlReportTo);
            this.boxReportTo.Items.Add(this.btnReportCustom);
            this.boxReportTo.Name = "boxReportTo";
            // 
            // ddlReportTo
            // 
            this.ddlReportTo.Label = "Report to:";
            this.ddlReportTo.Name = "ddlReportTo";
            // 
            // btnReportCustom
            // 
            this.btnReportCustom.Label = "Send";
            this.btnReportCustom.Name = "btnReportCustom";
            this.btnReportCustom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReportCustom_Click);
            // 
            // grSource
            // 
            this.grSource.Items.Add(this.btnSafeView);
            this.grSource.Items.Add(this.btnCopyToClipboard);
            this.grSource.Label = "Source";
            this.grSource.Name = "grSource";
            // 
            // btnSafeView
            // 
            this.btnSafeView.Image = global::SpamGrabber.Properties.Resources.search4doc;
            this.btnSafeView.Label = "Safe View";
            this.btnSafeView.Name = "btnSafeView";
            this.btnSafeView.ShowImage = true;
            this.btnSafeView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSafeView_Click);
            // 
            // btnCopyToClipboard
            // 
            this.btnCopyToClipboard.Image = global::SpamGrabber.Properties.Resources.spamgrab_copy;
            this.btnCopyToClipboard.Label = "Copy to Clipboard";
            this.btnCopyToClipboard.Name = "btnCopyToClipboard";
            this.btnCopyToClipboard.ShowImage = true;
            this.btnCopyToClipboard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyToClipboard_Click);
            // 
            // gpSettings
            // 
            this.gpSettings.Items.Add(this.btnSettings);
            this.gpSettings.Label = "Settings";
            this.gpSettings.Name = "gpSettings";
            // 
            // btnSettings
            // 
            this.btnSettings.Image = global::SpamGrabber.Properties.Resources.spamgrab_settings;
            this.btnSettings.Label = "Change Settings";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            // 
            // SpamGrabber_Ribbon
            // 
            this.Name = "SpamGrabber_Ribbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SpamGrabber_Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Report.ResumeLayout(false);
            this.Report.PerformLayout();
            this.ButtonGroup1.ResumeLayout(false);
            this.ButtonGroup1.PerformLayout();
            this.boxReportTo.ResumeLayout(false);
            this.boxReportTo.PerformLayout();
            this.grSource.ResumeLayout(false);
            this.grSource.PerformLayout();
            this.gpSettings.ResumeLayout(false);
            this.gpSettings.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Report;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup ButtonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportDefaultSpam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportDefaultHam;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxReportTo;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlReportTo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportCustom;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grSource;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSafeView;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyToClipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
    }

    partial class ThisRibbonCollection
    {
        internal SpamGrabber_Ribbon SpamGrabber_Ribbon
        {
            get { return this.GetRibbon<SpamGrabber_Ribbon>(); }
        }
    }
}
