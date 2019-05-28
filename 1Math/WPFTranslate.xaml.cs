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
            List<string> AcceptLanguages = new List<string>();
            foreach (string item in Translator.TranslatableLanguages.Keys)
            {
                AcceptLanguages.Add(Translator.TranslatableLanguages[item]["nativeName"]);
            }
            _fromLanguages.AddRange(AcceptLanguages);
            _toLanguages.AddRange(AcceptLanguages);
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                ComboBoxFromLanguage.SelectedItem = "自动检测";
                ComboBoxToLanguage.SelectedItem = "简体中文";
            }));
        }

        private void ButtonStartTranslate_Click(object sender, RoutedEventArgs e)
        {
            string toLanguageNativeName = (string)this.ComboBoxToLanguage.SelectedItem;
            string toLanguageCode=string.Empty;
            foreach (string code in Translator.TranslatableLanguages.Keys)
            {
                if (Translator.TranslatableLanguages[code]["nativeName"]==toLanguageNativeName)
                {
                    toLanguageCode = code;
                    break;
                }
            }
            Translation translation = new Translation(toLanguageCode);
            translation.ProgressChange += Translation_ProgressChange;
            translation.Start();
        }

        private void Translation_ProgressChange(object Sender, ProgressEventArgs progressEventArgs)
        {
            this.Dispatcher.BeginInvoke(new Action(()=> { ProgressBarForTranslation.Value = 100*progressEventArgs.NewProgress; }));
        }
    }
}
