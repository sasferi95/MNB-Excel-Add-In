namespace MNB_Excel_Add_In
{
    partial class MNBRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MNBRibbon()
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
            this.MNBGroup = this.Factory.CreateRibbonGroup();
            this.mnbDataBTN = this.Factory.CreateRibbonButton();
            this.logBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.MNBGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.MNBGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // MNBGroup
            // 
            this.MNBGroup.Items.Add(this.mnbDataBTN);
            this.MNBGroup.Items.Add(this.logBtn);
            this.MNBGroup.Label = "Sas Ferenc";
            this.MNBGroup.Name = "MNBGroup";
            // 
            // mnbDataBTN
            // 
            this.mnbDataBTN.Label = "MNB adatletöltés";
            this.mnbDataBTN.Name = "mnbDataBTN";
            this.mnbDataBTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mnbDataBTN_Click);
            // 
            // logBtn
            // 
            this.logBtn.Label = "Log";
            this.logBtn.Name = "logBtn";
            this.logBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.logBtn_Click);
            // 
            // MNBRibbon
            // 
            this.Name = "MNBRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MNBRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.MNBGroup.ResumeLayout(false);
            this.MNBGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup MNBGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mnbDataBTN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton logBtn;
    }

    partial class ThisRibbonCollection
    {
        internal MNBRibbon MNBRibbon
        {
            get { return this.GetRibbon<MNBRibbon>(); }
        }
    }
}
