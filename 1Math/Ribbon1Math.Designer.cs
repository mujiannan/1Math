using System;
using Microsoft.Office.Tools.Ribbon;

namespace _1Math
{
    partial class Ribbon1Math : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1Math()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.Tab1Math = this.Factory.CreateRibbonTab();
            this.GroupNet = this.Factory.CreateRibbonGroup();
            this.ButtonUrlCheck = this.Factory.CreateRibbonButton();
            this.splitButtonMediaDuration = this.Factory.CreateRibbonSplitButton();
            this.GroupDataCleaner = this.Factory.CreateRibbonGroup();
            this.ButtonAntiMerge = this.Factory.CreateRibbonButton();
            this.GroupText = this.Factory.CreateRibbonGroup();
            this.ButtonToEnglish = this.Factory.CreateRibbonSplitButton();
            this.ButtonTranslate = this.Factory.CreateRibbonButton();
            this.buttonQR = this.Factory.CreateRibbonButton();
            this.GroupOffSet = this.Factory.CreateRibbonGroup();
            this.ToggleButtonAutoOffSet = this.Factory.CreateRibbonToggleButton();
            this.BoxOffSet = this.Factory.CreateRibbonBox();
            this.DropDownOffSet = this.Factory.CreateRibbonDropDown();
            this.editBoxFactor = this.Factory.CreateRibbonEditBox();
            this.buttonMediaCheckSet = this.Factory.CreateRibbonButton();
            this.Tab1Math.SuspendLayout();
            this.GroupNet.SuspendLayout();
            this.GroupDataCleaner.SuspendLayout();
            this.GroupText.SuspendLayout();
            this.GroupOffSet.SuspendLayout();
            this.BoxOffSet.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1Math
            // 
            this.Tab1Math.Groups.Add(this.GroupNet);
            this.Tab1Math.Groups.Add(this.GroupDataCleaner);
            this.Tab1Math.Groups.Add(this.GroupText);
            this.Tab1Math.Groups.Add(this.GroupOffSet);
            this.Tab1Math.Label = "1Math";
            this.Tab1Math.Name = "Tab1Math";
            this.Tab1Math.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            // 
            // GroupNet
            // 
            this.GroupNet.Items.Add(this.ButtonUrlCheck);
            this.GroupNet.Items.Add(this.splitButtonMediaDuration);
            this.GroupNet.Label = "网络";
            this.GroupNet.Name = "GroupNet";
            // 
            // ButtonUrlCheck
            // 
            this.ButtonUrlCheck.Image = global::_1Math.Properties.Resources.链接1;
            this.ButtonUrlCheck.Label = "链接有效性";
            this.ButtonUrlCheck.Name = "ButtonUrlCheck";
            this.ButtonUrlCheck.ShowImage = true;
            this.ButtonUrlCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonUrlCheck_ClickAsync);
            // 
            // splitButtonMediaDuration
            // 
            this.splitButtonMediaDuration.Image = global::_1Math.Properties.Resources.秒表;
            this.splitButtonMediaDuration.Items.Add(this.buttonMediaCheckSet);
            this.splitButtonMediaDuration.Label = "媒体时长";
            this.splitButtonMediaDuration.Name = "splitButtonMediaDuration";
            this.splitButtonMediaDuration.ScreenTip = "检测媒体时长";
            this.splitButtonMediaDuration.SuperTip = "先选中一块连续区域（包含有效媒体链接），点击后，将在设置的偏移位置显示媒体时长，单位：（秒）";
            this.splitButtonMediaDuration.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SplitButtonMediaDurationAsync_Click);
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
            this.ButtonAntiMerge.ScreenTip = "取消合并单元格";
            this.ButtonAntiMerge.SuperTip = "批量取消选取中的合并单元格，并相对安全地自动填充。如果你只选中了一个单元格，那么会默认处理整个工作表。";
            this.ButtonAntiMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAntiMerge_ClickAsync);
            // 
            // GroupText
            // 
            this.GroupText.Items.Add(this.ButtonToEnglish);
            this.GroupText.Items.Add(this.buttonQR);
            this.GroupText.Label = "文本处理";
            this.GroupText.Name = "GroupText";
            // 
            // ButtonToEnglish
            // 
            this.ButtonToEnglish.Items.Add(this.ButtonTranslate);
            this.ButtonToEnglish.Label = "中译英";
            this.ButtonToEnglish.Name = "ButtonToEnglish";
            this.ButtonToEnglish.ScreenTip = "批量中译英";
            this.ButtonToEnglish.SuperTip = "选中具备有效文本的连续单元格，译文将显示在选区右侧";
            this.ButtonToEnglish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonToEnglish_ClickAsync);
            // 
            // ButtonTranslate
            // 
            this.ButtonTranslate.Label = "翻译...";
            this.ButtonTranslate.Name = "ButtonTranslate";
            this.ButtonTranslate.ScreenTip = "进行有更多详细设置的批量翻译";
            this.ButtonTranslate.ShowImage = true;
            this.ButtonTranslate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonTranslate_Click);
            // 
            // buttonQR
            // 
            this.buttonQR.Label = "生成二维码";
            this.buttonQR.Name = "buttonQR";
            this.buttonQR.ScreenTip = "批量生成二维码";
            this.buttonQR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonQRAsync_Click);
            // 
            // GroupOffSet
            // 
            this.GroupOffSet.Items.Add(this.ToggleButtonAutoOffSet);
            this.GroupOffSet.Items.Add(this.BoxOffSet);
            this.GroupOffSet.Label = "设置";
            this.GroupOffSet.Name = "GroupOffSet";
            // 
            // ToggleButtonAutoOffSet
            // 
            this.ToggleButtonAutoOffSet.Label = "自动输出偏移：右1*n";
            this.ToggleButtonAutoOffSet.Name = "ToggleButtonAutoOffSet";
            this.ToggleButtonAutoOffSet.ScreenTip = "指示数据输出的位置";
            this.ToggleButtonAutoOffSet.ShowImage = true;
            this.ToggleButtonAutoOffSet.SuperTip = "在工作表上读取数据并回写Excel时的“结果”相对于“数据源”的偏移量";
            this.ToggleButtonAutoOffSet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleButtonAutoOffSet_Click);
            // 
            // BoxOffSet
            // 
            this.BoxOffSet.Items.Add(this.DropDownOffSet);
            this.BoxOffSet.Items.Add(this.editBoxFactor);
            this.BoxOffSet.Name = "BoxOffSet";
            this.BoxOffSet.Visible = false;
            // 
            // DropDownOffSet
            // 
            ribbonDropDownItemImpl1.Label = "右";
            ribbonDropDownItemImpl2.Label = "左";
            this.DropDownOffSet.Items.Add(ribbonDropDownItemImpl1);
            this.DropDownOffSet.Items.Add(ribbonDropDownItemImpl2);
            this.DropDownOffSet.Label = "方向：";
            this.DropDownOffSet.Name = "DropDownOffSet";
            this.DropDownOffSet.ScreenTip = "偏移方向（左右）";
            this.DropDownOffSet.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDownOffSet_SelectionChanged);
            // 
            // editBoxFactor
            // 
            this.editBoxFactor.Label = "倍率：";
            this.editBoxFactor.Name = "editBoxFactor";
            this.editBoxFactor.ScreenTip = "偏移倍率值（数字）";
            this.editBoxFactor.SuperTip = "将以源数据的列数乘以次数值作为偏移量，例如，源数据共三列，你在此处输入“2”，则偏移量为2×3=6.";
            this.editBoxFactor.Text = "2";
            this.editBoxFactor.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditBoxFactor_TextChanged);
            // 
            // buttonMediaCheckSet
            // 
            this.buttonMediaCheckSet.Label = "媒体信息检测";
            this.buttonMediaCheckSet.Name = "buttonMediaCheckSet";
            this.buttonMediaCheckSet.ScreenTip = "更多设置";
            this.buttonMediaCheckSet.ShowImage = true;
            this.buttonMediaCheckSet.SuperTip = "可以获取更多媒体信息，包括时长、是否具有视频、是否有音频、分辨率";
            // 
            // Ribbon1Math
            // 
            this.Name = "Ribbon1Math";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Tab1Math);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.Tab1Math.ResumeLayout(false);
            this.Tab1Math.PerformLayout();
            this.GroupNet.ResumeLayout(false);
            this.GroupNet.PerformLayout();
            this.GroupDataCleaner.ResumeLayout(false);
            this.GroupDataCleaner.PerformLayout();
            this.GroupText.ResumeLayout(false);
            this.GroupText.PerformLayout();
            this.GroupOffSet.ResumeLayout(false);
            this.GroupOffSet.PerformLayout();
            this.BoxOffSet.ResumeLayout(false);
            this.BoxOffSet.PerformLayout();
            this.ResumeLayout(false);

        }






        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1Math;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupNet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonUrlCheck;
        internal RibbonGroup GroupDataCleaner;
        internal RibbonButton ButtonAntiMerge;
        internal RibbonGroup GroupText;
        internal RibbonSplitButton ButtonToEnglish;
        internal RibbonButton ButtonTranslate;
        internal RibbonGroup GroupOffSet;
        internal RibbonToggleButton ToggleButtonAutoOffSet;
        internal RibbonDropDown DropDownOffSet;
        internal RibbonBox BoxOffSet;
        internal RibbonEditBox editBoxFactor;
        internal RibbonButton buttonQR;
        internal RibbonSplitButton splitButtonMediaDuration;
        internal RibbonButton buttonMediaCheckSet;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1Math Ribbon1
        {
            get { return this.GetRibbon<Ribbon1Math>(); }
        }
    }
}
