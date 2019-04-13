using System.Net;
using System.Net.Http;
using System;
using System.Threading.Tasks;
using WMPLib;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
namespace _1Math
{
    public class CommonExcel
    {
        public static Excel.Application ExApp = Globals.ThisAddIn.Application;
        public static Excel.Window window = Globals.ThisAddIn.Application.ActiveWindow;
        public static Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
        public Excel.Range SelectedRange;
        public int x,y,m, n;
        public Excel.Range WriteInRange;
        public CommonExcel()
        {
            //ExApp.EnableEvents = false;//每次结束时别忘了激活事件监控
            SelectedRange = GetSelection();
            x = SelectedRange.Row;
            y = SelectedRange.Column;
            m = SelectedRange.Rows.Count;
            n = SelectedRange.Columns.Count;
            WriteInRange = SelectedRange.Offset[0, n];
            window.ScrollRow = x;
        }

        private Excel.Range GetSelection()
        {
            return ExApp.Application.Selection;
        }
    }
    public class Test
    {
        public delegate string ProcessDelegate(Object Rng);

        public void TestIt()
        {
            CommonExcel CR = new CommonExcel();
            for (int i = 0; i < 20; i++)
            {
                System.Threading.Thread.Sleep(1000);
                CR.SelectedRange.Value = i;
                System.Windows.Forms.Application.DoEvents();
            }
            //CR.SelectedRange.Cells[1, 1].value = Process(CR, new ProcessDelegate(Process1));
            //CR.SelectedRange.Cells[CR.SelectedRange.Rows.Count, CR.SelectedRange.Columns.Count].value = Process(CR, new ProcessDelegate(Process2));
        }
        public string Process(object Rng, ProcessDelegate doggy)
        {
            return (doggy(Rng));
        }
        public string Process1(object Rng)
        {
            return ("这里是程序1");
        }
        public string Process2(object Rng)
        {
            return (Rng.GetType().ToString());
        }
    }
    class Video
    {
        public Url url;
        public string Path;
        public double Duration
        {
            get
            {
                MyMediaPlayer myMediaPlayer = new MyMediaPlayer();
                return (myMediaPlayer.GetDuration(url.Str));
            }
        }
    }
    class Url
    {
        public new string ToString()
        {
            return Str;
        }
        private enum CheckStatus
        {
            Null, Checking, Checked
        }
        private Task checkTask;
        private CheckStatus checkStatus;
        public void CheckAccessibility()
        {
            checkTask = Check();
        }
        private async Task Check()
        {
            checkStatus = CheckStatus.Checking;
            HttpClient checkClient = new HttpClient();
            HttpResponseMessage httpResponseMessage = await checkClient.GetAsync(str, HttpCompletionOption.ResponseHeadersRead);
            accessibility = httpResponseMessage.IsSuccessStatusCode;
            checkClient.Dispose();
            checkStatus = CheckStatus.Checked;
        }
        private bool accessibility;
        public bool Accessibility
        {
            get
            {
                switch (checkStatus)
                {
                    case CheckStatus.Null:
                        CheckAccessibility();
                        checkTask.Wait();
                        break;
                    case CheckStatus.Checking:
                        checkTask.Wait();
                        break;
                    case CheckStatus.Checked:
                        break;
                    default:
                        break;
                }
                return accessibility;
            }
        }
        private string str;
        public string Str
        {
            get
            {
                return str;
            }
            set
            {
                if (value != str)
                {
                    checkStatus = CheckStatus.Null;
                    str = value;
                }

            }
        }
    }
    class MyMediaPlayer
    {
        WindowsMediaPlayer mediaPlayer;
        private string mediaUrl;
        private AutoResetEvent IsOpened = new AutoResetEvent(false);
        public MyMediaPlayer()
        {
            mediaPlayer = new WindowsMediaPlayer();
            mediaPlayer.OpenStateChange += MediaPlayer_OpenStateChange;
        }
        public double GetDuration(string url)
        {
            mediaUrl = url;
            Thread PlayThread = new Thread(Play);
            PlayThread.Start();
            IsOpened.WaitOne();
            double Duration = mediaPlayer.currentMedia.duration;
            mediaPlayer.controls.stop();
            mediaPlayer.URL = string.Empty;
            return (Duration);
        }
        private void Play()
        {
            mediaPlayer.URL = mediaUrl;
        }

        private void MediaPlayer_OpenStateChange(int NewState)
        {
            if (NewState == (int)WMPOpenState.wmposMediaOpen)
            {
                IsOpened.Set();
            }
        }
    }

    public class NetTask
    {
        private EventHandler Shutdown;
        private EventHandler Startup;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }
        public void CheckUrlsAccessibility()//专用于检查乂学的视频链接有效性
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            CommonExcel CE = new CommonExcel();
            string[,] Urls = new string[CE.m, CE.n];
            for (int i = 0; i < Urls.GetLength(0); i++)
            {
                for (int j = 0; j < Urls.GetLength(1); j++)
                {
                    Urls[i, j] = CE.SelectedRange.Cells[i+1, j+1].value;
                }
            }//读入数组，可以按列读
            bool[,] Accessibilities = new bool[CE.m, CE.n];
            Url url = new Url();
            int sum = 0;
            int t = 0;
            for (int i = 0; i < Urls.GetLength(0); i++)
            {
                for (int j = 0; j < Urls.GetLength(1); j++)
                {
                    url.Str =Urls[i,j];
                    Accessibilities[i, j] = url.Accessibility;
                    if (!url.Accessibility)
                    {
                        t++;
                    }
                    sum++;
                }
            }
            CE.WriteInRange.Value = Accessibilities;
            stopwatch.Stop();
            System.Windows.Forms.MessageBox.Show(@"耗时" + stopwatch.Elapsed.TotalSeconds + "秒，"
                                                    + "完成了" + sum + "个链接的有效性验证，其中" + t + "个无效");
        }
        public void CheckVideosLength()
        {
            CommonExcel CE = new CommonExcel();
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            int t = 0;
            MyMediaPlayer myMediaPlayer = new MyMediaPlayer();
            Url url = new Url();
            for (int i = 1; i <= CE.m; i++)
            {
                if (i > 8)
                {
                    CommonExcel.window.SmallScroll(1);//舒适地滚动
                }
                for (int j = 1; j <= CE.n; j++)
                {
                    url.Str = CE.SelectedRange[i, j].value;
                    CE.SelectedRange.Offset[0, 2 * CE.n].Cells[i, j].value = "正在检测时长……";//实测轻微降低了性能，占比很小
                    CE.SelectedRange.Offset[0, 2 * CE.n].Cells[i, j].value = myMediaPlayer.GetDuration(url.Str);
                    t++;
                }
            }
            stopwatch.Stop();
            System.Windows.Forms.MessageBox.Show(@"耗时" + stopwatch.Elapsed.TotalSeconds + "秒，" +
                                                "共选中了"+CE.SelectedRange.Count+"个单元格，"+
                                                "成功完成了" + t + "个视频时长的检测");
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

    }
}
