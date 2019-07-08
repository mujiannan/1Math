using AzureCognitiveTranslator;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            cancellationTokenSource.Token.Register(() => ExcelStatic.EndTask());
            StatusForm statusForm = new StatusForm();
            statusForm.Show();
            statusForm.FormClosing += (object s, FormClosingEventArgs eventArgs) => cancellationTokenSource.Cancel();
            AccessibilityChecker accessibilityChecker = new AccessibilityChecker()
            {
                ResultOffSet = this.OffSet
            };
            accessibilityChecker.Reportor.ProgressChange += statusForm.ChangeProgress;
            accessibilityChecker.Reportor.MessageChange += statusForm.ChangeMessage;
            try
            {
                await accessibilityChecker.StartAsync(cancellationTokenSource.Token);
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
                mergeAreas.Reportor.MessageChange += statusForm.ChangeMessage;
                mergeAreas.Reportor.ProgressChange += statusForm.ChangeProgress;
                CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
                cancellationTokenSource.Token.Register(() => ExcelStatic.EndTask());
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
        private async void ButtonToEnglish_ClickAsync(object sender, RibbonControlEventArgs e)
        {

            ExcelStatic.StartTask();
            try
            {
                
                Translator translator = new Translator(Properties.Resources.AzureCognitiveBaseUrl, Properties.Resources.AzureCognitiveKey);
                await Controller.TranslateSelectionAsync("en", translator);
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
            FormTranslator formWPF = new FormTranslator();
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
                        this.editBoxFactor.Text = "1";
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
            ExcelStatic.ResultOffset = OffSet;
        }

        private async void ButtonQRAsync_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelStatic.StartTask();
            string selectedPath;
            using (System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "请选择二维码的保存位置：";
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    selectedPath = folderBrowserDialog.SelectedPath;
                }
                else
                {
                    return;
                }
            }
            QRGenerator qRGenerator = new QRGenerator(selectedPath)
            {
                ResultOffSet = this.OffSet
            };
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            cancellationTokenSource.Token.Register(() => ExcelStatic.EndTask());
            StatusForm statusForm = new StatusForm();
            statusForm.FormClosing += (object s,System.Windows.Forms.FormClosingEventArgs formClosingEventArgs) => cancellationTokenSource.Cancel();
            qRGenerator.Reportor.MessageChange += statusForm.ChangeMessage;
            qRGenerator.Reportor.ProgressChange += statusForm.ChangeProgress;
            try
            {
                await qRGenerator.StartAsync(cancellationTokenSource.Token);
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

        private async void SplitButtonMediaDuration_Click(object sender, RibbonControlEventArgs e)
        {
            //我有点搞不懂了
            //如果这个事件处理程序作为Async void方法，在Ribbon里直接运行，就会有大量内存无法释放
            //但改成普通void方法，在方法内部用action封装一个async方法运行，就不会出现内存无法释放的情况
            //Action action = async() =>
            //  {
            ExcelStatic.StartTask();
            MediaInfoChecker mediaDurationChecker = new MediaInfoChecker()
            {
                ResultOffSet = this.OffSet,
                CheckDuration = true
            };
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            StatusForm statusForm = new StatusForm();
            statusForm.Show();
            mediaDurationChecker.Reportor.MessageChange += statusForm.ChangeMessage;
            mediaDurationChecker.Reportor.ProgressChange += statusForm.ChangeProgress;
            statusForm.FormClosing += (object s, System.Windows.Forms.FormClosingEventArgs formClosingEventArgs) => cancellationTokenSource.Cancel();
            cancellationTokenSource.Token.Register(() => ExcelStatic.EndTask());
            try
            {
                await mediaDurationChecker.StartAsync(cancellationTokenSource.Token);
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
            }
            finally
            {
                mediaDurationChecker.Reportor.MessageChange -= statusForm.ChangeMessage;
                mediaDurationChecker.Reportor.ProgressChange -= statusForm.ChangeProgress;
                mediaDurationChecker = null;
                System.GC.Collect();
                ExcelStatic.EndTask();
            }
            //  };
            //action.Invoke();

        }

        private void ButtonMoreMediaInfo_Click(object sender, RibbonControlEventArgs e)
        {
            FormCheckMediaInfo formCheckMediaInfo = new FormCheckMediaInfo();
            formCheckMediaInfo.Show();
        }
    }
}
