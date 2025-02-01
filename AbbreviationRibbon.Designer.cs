namespace AbbreviationWordAddin
{
    partial class AbbreviationRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AbbreviationRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnEnable = this.Factory.CreateRibbonButton();
            this.btnDisable = this.Factory.CreateRibbonButton();
            this.btnReplaceAll = this.Factory.CreateRibbonButton();
            this.btnHighlightAll = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnEnable);
            this.group1.Items.Add(this.btnDisable);
            this.group1.Items.Add(this.btnReplaceAll);
            this.group1.Items.Add(this.btnHighlightAll);
            this.group1.Label = "Abbreviation Tools";
            this.group1.Name = "group1";
            // 
            // btnEnable
            // 
            this.btnEnable.Label = "Enable Abbreviation";
            this.btnEnable.Name = "btnEnable";
            this.btnEnable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEnable_Click);
            // 
            // btnDisable
            // 
            this.btnDisable.Label = "Disable Abbreviation";
            this.btnDisable.Name = "btnDisable";
            this.btnDisable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisable_Click);
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.Label = "Replace All";
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceAll_Click);
            // 
            // btnHighlightAll
            // 
            this.btnHighlightAll.Label = "Highlight All";
            this.btnHighlightAll.Name = "btnHighlightAll";
            this.btnHighlightAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightAll_Click);
            // 
            // AbbreviationRibbon
            // 
            this.Name = "AbbreviationRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AbbreviationRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEnable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightAll;
    }

    partial class ThisRibbonCollection
    {
        internal AbbreviationRibbon AbbreviationRibbon
        {
            get { return this.GetRibbon<AbbreviationRibbon>(); }
        }
    }
}
