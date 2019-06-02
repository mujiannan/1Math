using AzureCognitiveTranslator;
using Microsoft.Office.Tools.Ribbon;
using System;

namespace _1Math
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void ButtonUrlCheck_Click(object sender, RibbonControlEventArgs e)
        {
            CE.StartTask();
            try
            {
                Accessibility accessibility = new Accessibility();
                BackGroundTask backGroundTask = new BackGroundTask(accessibility);
                backGroundTask.Start(new CancelableMethod(accessibility.Start));
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);

            }
            finally
            {
                CE.EndTask();
            }
        }
        private void ButtonAntiMerge_Click(object sender, RibbonControlEventArgs e)
        {
            CE.StartTask();
            try
            {
                MergeAreas mergeAreas = new MergeAreas();
                BackGroundTask backGroundTask = new BackGroundTask(mergeAreas);
                backGroundTask.Start(new CancelableMethod(mergeAreas.SafelyUnMergeAndFill));
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
                CE.EndTask();
            }
        }

        private void ButtonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            CE.StartTask();
            try
            {
                VideoLength videoLength = new VideoLength();
                BackGroundTask backGroundTask = new BackGroundTask(videoLength);
                backGroundTask.Start(new CancelableMethod(videoLength.Start));
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message);
            }
            finally
            {
                CE.EndTask();
            }

        }
        private async void ButtonToEnglish_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            CE.StartTask();
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
                CE.EndTask();
            }

        }
        private void ButtonTranslate_Click(object sender, RibbonControlEventArgs e)
        {
            FormWPF formWPF = new FormWPF();
            formWPF.Show();
        }
    }
}
