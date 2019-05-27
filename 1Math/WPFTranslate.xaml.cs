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

        private void Test_Click(object sender, RoutedEventArgs e)
        {
            this.ComboBoxFromLanguage.Items.Add("哼哼");
            this.ComboBoxFromLanguage.Items.Add("哈哈");
        }
    }
}
