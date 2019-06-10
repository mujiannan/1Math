using AzureCognitiveTranslator;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;

namespace _1Math
{
    public partial class Ribbon1Math
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private async void ButtonUrlCheck_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            ExcelStatic.StartTask();
            try
            {
                ExcelConcurrentTask excelConcurrent = new ExcelConcurrentTask
                                                                            (
                                                                                new AccessibilityChecker()
                                                                                {
                                                                                    ResultOffSet = this.OffSet
                                                                                }
                                                                            );
                await excelConcurrent.StartAsync();
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
            }
            finally
            {
                ExcelStatic.EndTask();
            }

        }
        private async void ButtonAntiMerge_ClickAsync(object sender, RibbonControlEventArgs e)
        {
#if DEBUG
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
#endif
            try
            {
                ExcelStatic.StartTask();
                MergeAreas mergeAreas = new MergeAreas();
                StatusForm statusForm = new StatusForm();
                statusForm.Show();
                mergeAreas.Reportor.MessageChange += statusForm.MessageLabel_TextChange;
                mergeAreas.Reportor.ProgressChange += statusForm.ProgressBar_ValueChange;
                CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
                statusForm.FormClosing += delegate
                {
                    cancellationTokenSource.Cancel();
                };
                await mergeAreas.SafelyUnMergeAndFill(cancellationTokenSource.Token);
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
            }
            finally
            {
                ExcelStatic.EndTask();
            }
#if DEBUG
            System.Windows.Forms.MessageBox.Show($"耗时{stopwatch.Elapsed.TotalSeconds.ToString()}秒");
#endif
        }

        private async void ButtonVideoLength_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            ExcelStatic.StartTask();
            try
            {
                ExcelConcurrentTask excelConcurrent = new ExcelConcurrentTask(new MediaDurationChecker() { ResultOffSet = this.OffSet });
                await excelConcurrent.StartAsync();
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
            }
            finally
            {
                ExcelStatic.EndTask();
            }
        }
        private async void ButtonToEnglish_ClickAsync(object sender, RibbonControlEventArgs e)
        {

            ExcelStatic.StartTask();
            try
            {
                Translator translator = new Translator(Properties.Resources.AzureCognitiveBaseUrl, Properties.Resources.AzureCognitiveKey);
                await Main.TranslateSelectionAsync("en", translator);
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
            }
            finally
            {
                ExcelStatic.EndTask();
            }

        }
        private void ButtonTranslate_Click(object sender, RibbonControlEventArgs e)
        {
            FormWPF formWPF = new FormWPF();
            formWPF.Show();
        }

        private void ToggleButtonAutoOffSet_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.ToggleButtonAutoOffSet.Checked)
            {
                this.ToggleButtonAutoOffSet.Label = "手动设置偏移量：";
                this.BoxOffSet.Visible = true;
            }
            else
            {
                this.ToggleButtonAutoOffSet.Label = "自动输出偏移：右1*n";
                this.BoxOffSet.Visible = false;
            }
            ExcelStatic.ResultOffset = OffSet;
        }
        public int[] OffSet
        {
            get
            {
                int k=1,x=1;
                if (this.ToggleButtonAutoOffSet.Checked)
                {
                    if (DropDownOffSet.SelectedItem.Label=="左")
                    {
                        k = -1;
                    }
                    else
                    {
                        k = 1;
                    }
                    bool xIsInt = int.TryParse(this.editBoxFactor.Text,out x);
                    if (!xIsInt)
                    {
                        x = 1;
                    }
                }
                return new int[2] { 0, k * x };
            }
        }
        private void DropDownOffSet_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            ExcelStatic.ResultOffset = OffSet;
        }

        private void EditBoxFactor_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int num;
            bool inputIsNum = int.TryParse(this.editBoxFactor.Text,out num);
            if (!inputIsNum)
            {
                this.editBoxFactor.Text = "";
            }
            ExcelStatic.ResultOffset = OffSet;
        }
    }
}
