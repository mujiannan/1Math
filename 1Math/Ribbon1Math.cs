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
                ExcelConcurrentTask excelConcurrent = new ExcelConcurrentTask(new AccessibilityChecker());
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
            try
            {
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
                ExcelConcurrentTask excelConcurrent = new ExcelConcurrentTask(new MediaDurationChecker());
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
    }
}
