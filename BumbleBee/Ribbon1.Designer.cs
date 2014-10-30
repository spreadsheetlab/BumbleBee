namespace ExcelAddIn3
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.BumbleBee = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.Preview = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.analyzeButton = this.Factory.CreateRibbonButton();
            this.selectSmellType = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.BumbleBee.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // BumbleBee
            // 
            this.BumbleBee.Groups.Add(this.group2);
            this.BumbleBee.Groups.Add(this.group1);
            this.BumbleBee.Label = "BumbleBee";
            this.BumbleBee.Name = "BumbleBee";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.dropDown1);
            this.group2.Items.Add(this.Preview);
            this.group2.Items.Add(this.separator1);
            this.group2.Items.Add(this.button2);
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button3);
            this.group2.Label = "Basic Options";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.Label = "Find applicable rewrites";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // dropDown1
            // 
            this.dropDown1.Label = "Rewrites possible";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.SizeString = "ietsminderlangesuperstring";
            this.dropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // Preview
            // 
            this.Preview.Label = "Preview";
            this.Preview.Name = "Preview";
            this.Preview.SizeString = "helelangesuperdeluzesuperlangsformulestring";
            this.Preview.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button2
            // 
            this.button2.Label = "Apply in Range";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button4
            // 
            this.button4.Label = "Apply in Sheet";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Label = "Apply Everywhere";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.analyzeButton);
            this.group1.Items.Add(this.selectSmellType);
            this.group1.Label = "Smells";
            this.group1.Name = "group1";
            // 
            // analyzeButton
            // 
            this.analyzeButton.Label = "Color smelly cells";
            this.analyzeButton.Name = "analyzeButton";
            this.analyzeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // selectSmellType
            // 
            this.selectSmellType.Enabled = false;
            this.selectSmellType.Label = "Show only";
            this.selectSmellType.Name = "selectSmellType";
            this.selectSmellType.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectSmellType_SelectionChanged);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.BumbleBee);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.BumbleBee.ResumeLayout(false);
            this.BumbleBee.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab BumbleBee;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Preview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton analyzeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown selectSmellType;

        #if DEBUG
            internal Microsoft.Office.Tools.Ribbon.RibbonButton Initialize;
        #endif

    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
