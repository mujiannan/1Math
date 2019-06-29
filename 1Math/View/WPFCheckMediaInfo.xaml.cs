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
using System.Threading;
using Correspondence;

namespace _1Math
{
    /// <summary>
    /// WPFCheckMediaInfo.xaml 的交互逻辑
    /// </summary>
    public partial class WPFCheckMediaInfo : UserControl
    {
        public WPFCheckMediaInfo()
        {
            InitializeComponent();
        }

        private async void CheckAsync_Click(object sender, RoutedEventArgs e)
        {
            ExcelStatic.StartTask();
            MediaInfoChecker mediaDurationChecker = new MediaInfoChecker()
            {
                ResultOffSet = ExcelStatic.ResultOffset,
                CheckDuration = this.CheckDuration.IsChecked.HasValue ? (bool)this.CheckDuration.IsChecked : false,
                CheckHasVideo = this.CheckHasVideo.IsChecked.HasValue ? (bool)this.CheckHasVideo.IsChecked : false,
                CheckHasAudio = this.CheckHasAudio.IsChecked.HasValue ? (bool)this.CheckHasAudio.IsChecked : false,
                CheckResolution = this.CheckResolution.IsChecked.HasValue ? (bool)this.CheckResolution.IsChecked : false
            };
            mediaDurationChecker.Reportor.ProgressChange += Reportor_ProgressChange;
            mediaDurationChecker.Reportor.MessageChange += Reportor_MessageChange;
            CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
            cancellationTokenSource.Token.Register(() => ExcelStatic.EndTask());
            this.Unloaded += (object s, System.Windows.RoutedEventArgs routedEventArgs) =>
            {
                cancellationTokenSource.Cancel();
            };
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
                ExcelStatic.EndTask();
            }
        }

        private void Reportor_MessageChange(object sender, Reportor.MessageEventArgs e)
        {
            try
            {
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.MessageForCheckMediaInfo.Text = e.NewMessage;
                }));
            }
            catch (Exception)
            {
            }
            
        }

        private void Reportor_ProgressChange(object sender, Reportor.ProgressEventArgs e)
        {
            try
            {
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.ProgressForCheckMediaInfo.Value =100*e.NewProgress;
                }));
            }
            catch (Exception)
            {
            }

        }
    }
}
