namespace OutlookFileDrag
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
            this.OutlookFileDragTab = this.Factory.CreateRibbonTab();
            this.options = this.Factory.CreateRibbonGroup();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.OutlookFileDragTab.SuspendLayout();
            this.options.SuspendLayout();
            this.SuspendLayout();
            // 
            // OutlookFileDragTab
            // 
            this.OutlookFileDragTab.Groups.Add(this.options);
            resources.ApplyResources(this.OutlookFileDragTab, "OutlookFileDragTab");
            this.OutlookFileDragTab.Name = "OutlookFileDragTab";
            // 
            // options
            // 
            this.options.Items.Add(this.comboBox1);
            resources.ApplyResources(this.options, "options");
            this.options.Name = "options";
            // 
            // comboBox1
            // 
            resources.ApplyResources(this.comboBox1, "comboBox1");
            this.comboBox1.Name = "comboBox1";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.OutlookFileDragTab);
            resources.ApplyResources(this, "$this");
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.OutlookFileDragTab.ResumeLayout(false);
            this.OutlookFileDragTab.PerformLayout();
            this.options.ResumeLayout(false);
            this.options.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab OutlookFileDragTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup options;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
