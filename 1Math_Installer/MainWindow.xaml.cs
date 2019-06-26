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

namespace _1Math_Installer
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            SetUp();
        }
        private async void SetUp()
        {
            Installer installer = new Installer();
            installer.Reportor.ProgressChange += (object sender, _1Math.Reportor.ProgressEventArgs e) => progressBar.Dispatcher.Invoke(()=>progressBar.Value = e.NewProgress);
            installer.Reportor.MessageChange += (object sender, _1Math.Reportor.MessageEventArgs e) =>messageTextBlock.Dispatcher.Invoke(()=>messageTextBlock.Text = e.NewMessage);
            try
            {
                await installer.StartInstallerAsync();
                this.Close();
            }
            catch (Exception ex)
            {
                messageTextBlock.Text = ex.Message;
            }
        }
    }
}
