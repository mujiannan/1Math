using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AzureCognitiveTranslator;
using Microsoft.Office.Tools.Ribbon;

namespace _1Math
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void ButtonUrlCheck_Click(object sender, RibbonControlEventArgs e)
        {
            Accessibility accessibility = new Accessibility();
            BackGroundTask backGroundTask = new BackGroundTask(accessibility);
            backGroundTask.Start(new CancelableMethod(accessibility.Start));
        }
        private void ButtonAntiMerge_Click(object sender, RibbonControlEventArgs e)
        {
            MergeAreas mergeAreas = new MergeAreas();
            BackGroundTask backGroundTask = new BackGroundTask(mergeAreas);
            backGroundTask.Start(new CancelableMethod(mergeAreas.SafelyUnMergeAndFill));
        }

        private void ButtonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            VideoLength videoLength = new VideoLength();
            BackGroundTask backGroundTask = new BackGroundTask(videoLength);
            backGroundTask.Start(new CancelableMethod(videoLength.Start));
        }
        private async void ButtonToEnglish_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            Translator translator = new Translator(Properties.Resources.AzureCognitiveBaseUrl, Properties.Resources.AzureCognitiveKey);
            await Main.TranslateSelectionAsync("en", translator);
        }
        private void ButtonTranslate_Click(object sender, RibbonControlEventArgs e)
        {
            FormWPF formWPF = new FormWPF();
            formWPF.Show();
        }
    }
}
