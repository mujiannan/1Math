using AzureCognitiveTranslator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

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
        private List<string> _fromLanguages = new List<string>();
        private List<string> _toLanguages = new List<string>();
        private void UserControl_Initialized(object sender, EventArgs e)
        {
            this.ComboBoxFromLanguage.ItemsSource = _fromLanguages;
            this.ComboBoxToLanguage.ItemsSource = _toLanguages;
            _fromLanguages.Add("自动检测");
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
            _fromLanguages.AddRange(AcceptLanguages);
            _toLanguages.AddRange(AcceptLanguages);
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                ComboBoxFromLanguage.SelectedItem = "自动检测";
                ComboBoxToLanguage.SelectedItem = "简体中文";
            }));
        }

        private async void ButtonStartTranslate_ClickAsync(object sender, RoutedEventArgs e)
        {
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
                this.TextBlockTime.Text = "耗时: " + CE.Elapse + "秒";
            }
            catch (Exception)
            {
                CE.EndTask();
            }

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
