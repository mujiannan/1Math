using System;
using Microsoft.Office.Tools.Ribbon;

namespace _1Math
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
            this.Tab1Math = this.Factory.CreateRibbonTab();
            this.GroupNet = this.Factory.CreateRibbonGroup();
            this.ButtonUrlCheck = this.Factory.CreateRibbonButton();
            this.buttonVideoLength = this.Factory.CreateRibbonButton();
            this.GroupDataCleaner = this.Factory.CreateRibbonGroup();
            this.ButtonAntiMerge = this.Factory.CreateRibbonButton();
            this.Tab1Math.SuspendLayout();
            this.GroupNet.SuspendLayout();
            this.GroupDataCleaner.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1Math
            // 
            this.Tab1Math.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1Math.Groups.Add(this.GroupNet);
            this.Tab1Math.Groups.Add(this.GroupDataCleaner);
            this.Tab1Math.Label = "1Math";
            this.Tab1Math.Name = "Tab1Math";
            this.Tab1Math.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            // 
            // GroupNet
            // 
            this.GroupNet.Items.Add(this.ButtonUrlCheck);
            this.GroupNet.Items.Add(this.buttonVideoLength);
            this.GroupNet.Label = "网络";
            this.GroupNet.Name = "GroupNet";
            // 
            // ButtonUrlCheck
            // 
            this.ButtonUrlCheck.Image = ((System.Drawing.Image)(resources.GetObject("ButtonUrlCheck.Image")));
            this.ButtonUrlCheck.Label = "链接有效性";
            this.ButtonUrlCheck.Name = "ButtonUrlCheck";
            this.ButtonUrlCheck.ShowImage = true;
            this.ButtonUrlCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonUrlCheck_Click);
            // 
            // buttonVideoLength
            // 
            this.buttonVideoLength.Image = ((System.Drawing.Image)(resources.GetObject("buttonVideoLength.Image")));
            this.buttonVideoLength.Label = "视频时长";
            this.buttonVideoLength.Name = "buttonVideoLength";
            this.buttonVideoLength.ShowImage = true;
            this.buttonVideoLength.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonVideoLength_Click);
            // 
            // GroupDataCleaner
            // 
            this.GroupDataCleaner.Items.Add(this.ButtonAntiMerge);
            this.GroupDataCleaner.Label = "数据清洗";
            this.GroupDataCleaner.Name = "GroupDataCleaner";
            // 
            // ButtonAntiMerge
            // 
            this.ButtonAntiMerge.Label = "取消合并";
            this.ButtonAntiMerge.Name = "ButtonAntiMerge";
            this.ButtonAntiMerge.ScreenTip = "批量取消合并单元格，并相对安全地自动填充";
            this.ButtonAntiMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAntiMerge_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Tab1Math);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.Tab1Math.ResumeLayout(false);
            this.Tab1Math.PerformLayout();
            this.GroupNet.ResumeLayout(false);
            this.GroupNet.PerformLayout();
            this.GroupDataCleaner.ResumeLayout(false);
            this.GroupDataCleaner.PerformLayout();
            this.ResumeLayout(false);

        }



        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1Math;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupNet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonUrlCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonVideoLength;
        internal RibbonGroup GroupDataCleaner;
        internal RibbonButton ButtonAntiMerge;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
