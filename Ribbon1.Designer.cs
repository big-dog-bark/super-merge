namespace SuperMerge
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tabSuperMerge = this.Factory.CreateRibbonTab();
            this.gpMergeAndExport = this.Factory.CreateRibbonGroup();
            this.btnExportDOCX = this.Factory.CreateRibbonButton();
            this.btnExportPDF = this.Factory.CreateRibbonButton();
            this.tabSuperMerge.SuspendLayout();
            this.gpMergeAndExport.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSuperMerge
            // 
            this.tabSuperMerge.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSuperMerge.Groups.Add(this.gpMergeAndExport);
            this.tabSuperMerge.Label = "SuperMerge";
            this.tabSuperMerge.Name = "tabSuperMerge";
            // 
            // gpMergeAndExport
            // 
            this.gpMergeAndExport.Items.Add(this.btnExportDOCX);
            this.gpMergeAndExport.Items.Add(this.btnExportPDF);
            this.gpMergeAndExport.Label = "合并后导出";
            this.gpMergeAndExport.Name = "gpMergeAndExport";
            // 
            // btnExportDOCX
            // 
            this.btnExportDOCX.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportDOCX.Image = ((System.Drawing.Image)(resources.GetObject("btnExportDOCX.Image")));
            this.btnExportDOCX.Label = "指定文件名模板导出DOCX";
            this.btnExportDOCX.Name = "btnExportDOCX";
            this.btnExportDOCX.ShowImage = true;
            this.btnExportDOCX.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportDOCX_Click);
            // 
            // btnExportPDF
            // 
            this.btnExportPDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportPDF.Image = ((System.Drawing.Image)(resources.GetObject("btnExportPDF.Image")));
            this.btnExportPDF.Label = "指定文件名模板导出为PDF";
            this.btnExportPDF.Name = "btnExportPDF";
            this.btnExportPDF.ShowImage = true;
            this.btnExportPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportPDF_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabSuperMerge);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabSuperMerge.ResumeLayout(false);
            this.tabSuperMerge.PerformLayout();
            this.gpMergeAndExport.ResumeLayout(false);
            this.gpMergeAndExport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSuperMerge;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpMergeAndExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportDOCX;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportPDF;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
