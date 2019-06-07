using AzureCognitiveTranslator;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.Diagnostics;

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
