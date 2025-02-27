namespace ol.clean
{
    partial class Cleaner : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Cleaner()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            this.tabClean = this.Factory.CreateRibbonTab();
            this.grpClean = this.Factory.CreateRibbonGroup();
            this.drpMarkRead = this.Factory.CreateRibbonDropDown();
            this.chkCommit = this.Factory.CreateRibbonCheckBox();
            this.grpAdd = this.Factory.CreateRibbonGroup();
            this.drpPeriod = this.Factory.CreateRibbonDropDown();
            this.grpFind = this.Factory.CreateRibbonGroup();
            this.btnLogFolder = this.Factory.CreateRibbonButton();
            this.btnClean = this.Factory.CreateRibbonButton();
            this.btnDelete = this.Factory.CreateRibbonButton();
            this.btnManageRules = this.Factory.CreateRibbonButton();
            this.btnAddDomain = this.Factory.CreateRibbonButton();
            this.btnAddExact = this.Factory.CreateRibbonButton();
            this.btnFindDomain = this.Factory.CreateRibbonButton();
            this.btnFindExact = this.Factory.CreateRibbonButton();
            this.tabClean.SuspendLayout();
            this.grpClean.SuspendLayout();
            this.grpAdd.SuspendLayout();
            this.grpFind.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabClean
            // 
            this.tabClean.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabClean.Groups.Add(this.grpClean);
            this.tabClean.Groups.Add(this.grpAdd);
            this.tabClean.Groups.Add(this.grpFind);
            this.tabClean.Label = "Clean";
            this.tabClean.Name = "tabClean";
            // 
            // grpClean
            // 
            this.grpClean.Items.Add(this.btnLogFolder);
            this.grpClean.Items.Add(this.drpMarkRead);
            this.grpClean.Items.Add(this.chkCommit);
            this.grpClean.Items.Add(this.btnClean);
            this.grpClean.Items.Add(this.btnDelete);
            this.grpClean.Label = "Clean";
            this.grpClean.Name = "grpClean";
            // 
            // drpMarkRead
            // 
            ribbonDropDownItemImpl1.Label = "Unchanged";
            ribbonDropDownItemImpl2.Label = "Mark read";
            ribbonDropDownItemImpl3.Label = "Mark unread";
            this.drpMarkRead.Items.Add(ribbonDropDownItemImpl1);
            this.drpMarkRead.Items.Add(ribbonDropDownItemImpl2);
            this.drpMarkRead.Items.Add(ribbonDropDownItemImpl3);
            this.drpMarkRead.Label = "Mark read";
            this.drpMarkRead.Name = "drpMarkRead";
            // 
            // chkCommit
            // 
            this.chkCommit.Description = "Log only or commit";
            this.chkCommit.Label = "Commit";
            this.chkCommit.Name = "chkCommit";
            // 
            // grpAdd
            // 
            this.grpAdd.Items.Add(this.btnManageRules);
            this.grpAdd.Items.Add(this.drpPeriod);
            this.grpAdd.Items.Add(this.btnAddDomain);
            this.grpAdd.Items.Add(this.btnAddExact);
            this.grpAdd.Label = "Add";
            this.grpAdd.Name = "grpAdd";
            // 
            // drpPeriod
            // 
            ribbonDropDownItemImpl4.Label = "1 week";
            ribbonDropDownItemImpl5.Label = "2 weeks";
            ribbonDropDownItemImpl6.Label = "3 weeks";
            ribbonDropDownItemImpl7.Label = "1 month";
            ribbonDropDownItemImpl8.Label = "2 months";
            ribbonDropDownItemImpl9.Label = "3 months";
            this.drpPeriod.Items.Add(ribbonDropDownItemImpl4);
            this.drpPeriod.Items.Add(ribbonDropDownItemImpl5);
            this.drpPeriod.Items.Add(ribbonDropDownItemImpl6);
            this.drpPeriod.Items.Add(ribbonDropDownItemImpl7);
            this.drpPeriod.Items.Add(ribbonDropDownItemImpl8);
            this.drpPeriod.Items.Add(ribbonDropDownItemImpl9);
            this.drpPeriod.Label = "Period";
            this.drpPeriod.Name = "drpPeriod";
            // 
            // grpFind
            // 
            this.grpFind.Items.Add(this.btnFindDomain);
            this.grpFind.Items.Add(this.btnFindExact);
            this.grpFind.Label = "Find";
            this.grpFind.Name = "grpFind";
            // 
            // btnLogFolder
            // 
            this.btnLogFolder.Description = "Open log folder";
            this.btnLogFolder.Image = global::ol.clean.Properties.Resources.Folder_6222;
            this.btnLogFolder.Label = "Log folder";
            this.btnLogFolder.Name = "btnLogFolder";
            this.btnLogFolder.ShowImage = true;
            this.btnLogFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogFolder_Click);
            // 
            // btnClean
            // 
            this.btnClean.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnClean.Description = "Clean now";
            this.btnClean.Image = global::ol.clean.Properties.Resources.DeadLetterMessages_5733_32;
            this.btnClean.Label = "Clean Now";
            this.btnClean.Name = "btnClean";
            this.btnClean.ShowImage = true;
            this.btnClean.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClean_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDelete.Description = "Delete";
            this.btnDelete.Image = global::ol.clean.Properties.Resources.delete;
            this.btnDelete.Label = "Delete";
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.ShowImage = true;
            this.btnDelete.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDelete_Click);
            // 
            // btnManageRules
            // 
            this.btnManageRules.Description = "Mange list of rules";
            this.btnManageRules.Image = global::ol.clean.Properties.Resources.Table_748;
            this.btnManageRules.Label = "Manage rules";
            this.btnManageRules.Name = "btnManageRules";
            this.btnManageRules.ShowImage = true;
            this.btnManageRules.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnManageRules_Click);
            // 
            // btnAddDomain
            // 
            this.btnAddDomain.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddDomain.Description = "Add Domain";
            this.btnAddDomain.Image = global::ol.clean.Properties.Resources.FilteredObject13400_128x128;
            this.btnAddDomain.Label = "Add Domain";
            this.btnAddDomain.Name = "btnAddDomain";
            this.btnAddDomain.ShowImage = true;
            this.btnAddDomain.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddDomain_Click);
            // 
            // btnAddExact
            // 
            this.btnAddExact.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddExact.Description = "Add Exact";
            this.btnAddExact.Image = global::ol.clean.Properties.Resources.ParametersInfo_2423_128x128;
            this.btnAddExact.Label = "Add Exact";
            this.btnAddExact.Name = "btnAddExact";
            this.btnAddExact.ShowImage = true;
            this.btnAddExact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddExact_Click);
            // 
            // btnFindDomain
            // 
            this.btnFindDomain.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFindDomain.Description = "Find Domain";
            this.btnFindDomain.Image = global::ol.clean.Properties.Resources.FilteredObject13400_128x128;
            this.btnFindDomain.Label = "Find Domain";
            this.btnFindDomain.Name = "btnFindDomain";
            this.btnFindDomain.ShowImage = true;
            this.btnFindDomain.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFindDomain_Click);
            // 
            // btnFindExact
            // 
            this.btnFindExact.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFindExact.Description = "Find Exact";
            this.btnFindExact.Image = global::ol.clean.Properties.Resources.ParametersInfo_2423_128x128;
            this.btnFindExact.Label = "Find Exact";
            this.btnFindExact.Name = "btnFindExact";
            this.btnFindExact.ShowImage = true;
            this.btnFindExact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFindExact_Click);
            // 
            // Cleaner
            // 
            this.Name = "Cleaner";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabClean);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Cleaner_Load);
            this.tabClean.ResumeLayout(false);
            this.tabClean.PerformLayout();
            this.grpClean.ResumeLayout(false);
            this.grpClean.PerformLayout();
            this.grpAdd.ResumeLayout(false);
            this.grpAdd.PerformLayout();
            this.grpFind.ResumeLayout(false);
            this.grpFind.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAdd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddDomain;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddExact;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpPeriod;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkCommit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnManageRules;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpMarkRead;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDelete;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFind;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFindDomain;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFindExact;
    }

    partial class ThisRibbonCollection
    {
        internal Cleaner Cleaner
        {
            get { return this.GetRibbon<Cleaner>(); }
        }
    }
}
