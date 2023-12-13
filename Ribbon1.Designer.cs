namespace OutlookAddIn_Report_Spam
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.HomeTab = this.Factory.CreateRibbonTab();
            this.ReportBtnGP = this.Factory.CreateRibbonGroup();
            this.ReportBtn = this.Factory.CreateRibbonButton();
            this.ReadMessageTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.HomeTab.SuspendLayout();
            this.ReportBtnGP.SuspendLayout();
            this.ReadMessageTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // HomeTab
            // 
            this.HomeTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.HomeTab.ControlId.OfficeId = "TabMail";
            this.HomeTab.Groups.Add(this.ReportBtnGP);
            this.HomeTab.Label = "TabMail";
            this.HomeTab.Name = "HomeTab";
            // 
            // ReportBtnGP
            // 
            this.ReportBtnGP.Items.Add(this.ReportBtn);
            this.ReportBtnGP.Label = "Report";
            this.ReportBtnGP.Name = "ReportBtnGP";
            // 
            // ReportBtn
            // 
            this.ReportBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ReportBtn.Image = global::OutlookAddIn_Report_Spam.Properties.Resources.Report_Icon;
            this.ReportBtn.Label = "Report Email";
            this.ReportBtn.Name = "ReportBtn";
            this.ReportBtn.ShowImage = true;
            this.ReportBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReportBtn_Click);
            // 
            // ReadMessageTab
            // 
            this.ReadMessageTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ReadMessageTab.ControlId.OfficeId = "TabReadMessage";
            this.ReadMessageTab.Groups.Add(this.group1);
            this.ReadMessageTab.Label = "TabReadMessage";
            this.ReadMessageTab.Name = "ReadMessageTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Report";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::OutlookAddIn_Report_Spam.Properties.Resources.Report_Icon;
            this.button1.Label = "Report Email";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReportBtn_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = resources.GetString("$this.RibbonType");
            this.Tabs.Add(this.HomeTab);
            this.Tabs.Add(this.ReadMessageTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.HomeTab.ResumeLayout(false);
            this.HomeTab.PerformLayout();
            this.ReportBtnGP.ResumeLayout(false);
            this.ReportBtnGP.PerformLayout();
            this.ReadMessageTab.ResumeLayout(false);
            this.ReadMessageTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab HomeTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ReportBtnGP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReportBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab ReadMessageTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
