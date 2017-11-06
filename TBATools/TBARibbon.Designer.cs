namespace TBATools
{
    partial class TBARibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TBARibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.TBAToolsTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnToColByCol = this.Factory.CreateRibbonButton();
            this.btnToColByRow = this.Factory.CreateRibbonButton();
            this.btnToRowByCol = this.Factory.CreateRibbonButton();
            this.btnToRowByRow = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.TBAToolsTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // TBAToolsTab
            // 
            this.TBAToolsTab.Groups.Add(this.group1);
            this.TBAToolsTab.Groups.Add(this.group2);
            this.TBAToolsTab.Label = "TBATools";
            this.TBAToolsTab.Name = "TBAToolsTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnToColByCol);
            this.group1.Items.Add(this.btnToColByRow);
            this.group1.Items.Add(this.btnToRowByCol);
            this.group1.Items.Add(this.btnToRowByRow);
            this.group1.Label = "矩阵操作";
            this.group1.Name = "group1";
            // 
            // btnToColByCol
            // 
            this.btnToColByCol.Image = global::TBATools.Properties.Resources.ToColByCol;
            this.btnToColByCol.Label = "转列（按列）";
            this.btnToColByCol.Name = "btnToColByCol";
            this.btnToColByCol.ShowImage = true;
            this.btnToColByCol.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToColByCol_Click);
            // 
            // btnToColByRow
            // 
            this.btnToColByRow.Image = global::TBATools.Properties.Resources.ToColByRow;
            this.btnToColByRow.Label = "转列（按行）";
            this.btnToColByRow.Name = "btnToColByRow";
            this.btnToColByRow.ShowImage = true;
            this.btnToColByRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToColByRow_Click);
            // 
            // btnToRowByCol
            // 
            this.btnToRowByCol.Image = global::TBATools.Properties.Resources.ToRowByCol;
            this.btnToRowByCol.Label = "转行（按列）";
            this.btnToRowByCol.Name = "btnToRowByCol";
            this.btnToRowByCol.ShowImage = true;
            this.btnToRowByCol.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToRowByCol_Click);
            // 
            // btnToRowByRow
            // 
            this.btnToRowByRow.Image = global::TBATools.Properties.Resources.ToRowByRow;
            this.btnToRowByRow.Label = "转行（按行）";
            this.btnToRowByRow.Name = "btnToRowByRow";
            this.btnToRowByRow.ShowImage = true;
            this.btnToRowByRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToRowByRow_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnAbout);
            this.group2.Label = "关于";
            this.group2.Name = "group2";
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "关于";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // TBARibbon
            // 
            this.Name = "TBARibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TBAToolsTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TBARibbon_Load);
            this.TBAToolsTab.ResumeLayout(false);
            this.TBAToolsTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TBAToolsTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToColByCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToColByRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToRowByCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToRowByRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
    }

    partial class ThisRibbonCollection
    {
        internal TBARibbon TBARibbon
        {
            get { return this.GetRibbon<TBARibbon>(); }
        }
    }
}
