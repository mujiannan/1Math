using AzureCognitiveTranslator;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
namespace _1Math
{
    /// <summary>
    /// WPFTranslate.xaml 的交互逻辑
    /// </summary>
    public partial class WPFTranslate : UserControl
    {
        public WPFTranslate()
        {
            InitializeComponent();
        }
        private List<string> _toLanguages = new List<string>();
        private void UserControl_Initialized(object sender, EventArgs e)
        {
            Task task = new Task(new Action(SetAcceptLanguages));
            task.Start();
        }
        private void SetAcceptLanguages()
        {
            Translator translator = new Translator(Properties.Resources.AzureCognitiveBaseUrl, Properties.Resources.AzureCognitiveKey);
            Dictionary<string, Translator.Language> translatableLanguages;
            try
            {
                translatableLanguages = translator.TranslatableLanguages;
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Can't get translableLanguage, perhaps exception on network");
                return;
            }
            List<string> AcceptLanguages = new List<string>();
            foreach (string code in translatableLanguages.Keys)
            {
                AcceptLanguages.Add(translatableLanguages[code].nativeName);
            }
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                ComboBoxFromLanguage.Items.Add("自动检测");
                ComboBoxFromLanguage.SelectedItem = "自动检测";
                foreach (string code in translatableLanguages.Keys)
                {
                    ComboBoxToLanguage.Items.Add(translatableLanguages[code].nativeName);
                }
                ComboBoxToLanguage.SelectedItem = "简体中文";
            }));
        }

        private async void ButtonStartTranslate_ClickAsync(object sender, RoutedEventArgs e)
        {
            CE.StartTask();
            try
            {
                Translator translator = new Translator(Properties.Resources.AzureCognitiveBaseUrl, Properties.Resources.AzureCognitiveKey);
                string toLanguageNativeName = (string)this.ComboBoxToLanguage.SelectedItem;
                string toLanguageCode = string.Empty;
                foreach (string code in translator.TranslatableLanguages.Keys)
                {
                    if (translator.TranslatableLanguages[code].nativeName == toLanguageNativeName)
                    {
                        toLanguageCode = code;
                        break;
                    }
                }
                translator.ProgressChange += Translator_ProgressChange;
                await Main.TranslateSelectionAsync(toLanguageCode, translator);
            }
            catch (Exception)
            {
                this.Dispatcher.Invoke(new Action(() => this.TextBlockTime.Text = "翻译失败"));
            }
            finally
            {
                CE.EndTask();
            }
            this.Dispatcher.Invoke(new Action(() => this.TextBlockTime.Text = "耗时: " + CE.Elapse + "秒"));

        }



        private void Translator_ProgressChange(object sender, Translator.TranslatingEventArgs translatingEventArgs)
        {
            try
            {
                this.Dispatcher.BeginInvoke(new Action(() => this.ProgressBarForTranslation.Value = 100 * translatingEventArgs.NewProgress));
            }
            catch (Exception)
            {

            }

        }
        private void Translation_ProgressChange(object Sender, ProgressEventArgs progressEventArgs)
        {
            try
            {
                this.Dispatcher.BeginInvoke(new Action(() => { ProgressBarForTranslation.Value = 100 * progressEventArgs.NewProgress; }));
            }
            catch (Exception)
            {

            }
        }
    }
}
