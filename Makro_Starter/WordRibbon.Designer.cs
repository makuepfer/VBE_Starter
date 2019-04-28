namespace Makro_Starter
{
    partial class WordRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WordRibbon()
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
            this.tabDevTools = this.Factory.CreateRibbonTab();
            this.grpVBA = this.Factory.CreateRibbonGroup();
            this.btnOpenVBE = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tabDevTools.SuspendLayout();
            this.grpVBA.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tabDevTools
            // 
            this.tabDevTools.Groups.Add(this.grpVBA);
            this.tabDevTools.Label = "DevTools";
            this.tabDevTools.Name = "tabDevTools";
            // 
            // grpVBA
            // 
            this.grpVBA.Items.Add(this.btnOpenVBE);
            this.grpVBA.Label = "VBA";
            this.grpVBA.Name = "grpVBA";
            // 
            // btnOpenVBE
            // 
            this.btnOpenVBE.Label = "Open VBE";
            this.btnOpenVBE.Name = "btnOpenVBE";
            this.btnOpenVBE.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenVBE_Click);
            // 
            // WordRibbon
            // 
            this.Name = "WordRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabDevTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WordRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabDevTools.ResumeLayout(false);
            this.tabDevTools.PerformLayout();
            this.grpVBA.ResumeLayout(false);
            this.grpVBA.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabDevTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpVBA;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenVBE;
    }

    partial class ThisRibbonCollection
    {
        internal WordRibbon WordRibbon
        {
            get { return this.GetRibbon<WordRibbon>(); }
        }
    }
}
