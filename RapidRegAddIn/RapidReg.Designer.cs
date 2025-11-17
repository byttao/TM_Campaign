

namespace RapidRegAddIn
{
    partial class RapidReg : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RapidReg()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RapidReg));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_Danpinbao = this.Factory.CreateRibbonButton();
            this.btn_Price = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.l_FolderPath = this.Factory.CreateRibbonLabel();
            this.btn_PathExplorer = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_Quick = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "活动报名";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_Danpinbao);
            this.group1.Items.Add(this.btn_Price);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "报名活动";
            this.group1.Name = "group1";
            // 
            // btn_Danpinbao
            // 
            this.btn_Danpinbao.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Danpinbao.Image = ((System.Drawing.Image)(resources.GetObject("btn_Danpinbao.Image")));
            this.btn_Danpinbao.Label = "单品宝处理";
            this.btn_Danpinbao.Name = "btn_Danpinbao";
            this.btn_Danpinbao.ShowImage = true;
            this.btn_Danpinbao.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Danpinbao_Click);
            // 
            // btn_Price
            // 
            this.btn_Price.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Price.Image = ((System.Drawing.Image)(resources.GetObject("btn_Price.Image")));
            this.btn_Price.Label = "计算活动价";
            this.btn_Price.Name = "btn_Price";
            this.btn_Price.ShowImage = true;
            this.btn_Price.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Price_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // box1
            // 
            this.box1.Items.Add(this.label1);
            this.box1.Items.Add(this.l_FolderPath);
            this.box1.Items.Add(this.btn_PathExplorer);
            this.box1.Name = "box1";
            // 
            // label1
            // 
            this.label1.Label = "活动路径：";
            this.label1.Name = "label1";
            // 
            // l_FolderPath
            // 
            this.l_FolderPath.Label = "未选中活动文件夹";
            this.l_FolderPath.Name = "l_FolderPath";
            // 
            // btn_PathExplorer
            // 
            this.btn_PathExplorer.Label = "...";
            this.btn_PathExplorer.Name = "btn_PathExplorer";
            this.btn_PathExplorer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_PathExplorer_Click);
            // 
            // button1
            // 
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_Quick);
            this.group2.Label = "快速处理";
            this.group2.Name = "group2";
            // 
            // btn_Quick
            // 
            this.btn_Quick.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Quick.Image = ((System.Drawing.Image)(resources.GetObject("btn_Quick.Image")));
            this.btn_Quick.Label = "表格快速处理";
            this.btn_Quick.Name = "btn_Quick";
            this.btn_Quick.ShowImage = true;
            this.btn_Quick.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Quick_Click);
            // 
            // RapidReg
            // 
            this.Name = "RapidReg";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RapidReg_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Danpinbao;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_PathExplorer;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel l_FolderPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Price;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Quick;
    }

    partial class ThisRibbonCollection
    {
        internal RapidReg RapidReg
        {
            get { return this.GetRibbon<RapidReg>(); }
        }
    }
}
