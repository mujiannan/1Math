using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Http;
using System;
using System.Threading.Tasks;
using WMPLib;
using System.Threading;
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
    public struct Video
    {
        public Url url;
        public string Path;
        public double Duration
        {
            get
            {
                MyMediaPlayer myMediaPlayer = new MyMediaPlayer();
                return (myMediaPlayer.GetDuration(url));
            }
        }
    }
    public struct Url
    {
        public string Value;
        public async Task<bool> CheckAccessibilityAsync()//用于快速返回链接内容
        {
            HttpClient HttpCheckTask = new HttpClient();
            //HttpCheckTask.CancelPendingRequests();
            //HttpCheckTask.Timeout =TimeSpan.FromSeconds(2);//不要乱设超时
            HttpResponseMessage httpResponse = await HttpCheckTask.GetAsync(Value, HttpCompletionOption.ResponseHeadersRead);
            bool Accessibility = (httpResponse.StatusCode == HttpStatusCode.OK);
            HttpCheckTask.Dispose();
            return Accessibility;
        }
    }
    public class MyMediaPlayer
    {
        WindowsMediaPlayer mediaPlayer;
        private string Url;
        private AutoResetEvent IsOpened = new AutoResetEvent(false);
        public double GetDuration(Url url)
        {
            Url = url.Value;
            Thread PlayThread = new Thread(PlayInNewPlayer);
            PlayThread.Start();
            IsOpened.WaitOne();
            double Duration = mediaPlayer.currentMedia.duration;
            mediaPlayer.close();
            return (Duration);
        }
        public double GetDuration(string url,ref WindowsMediaPlayer player)
        {
            mediaPlayer = player;
            Url = url;
            Thread PlayThread = new Thread(Play);
            PlayThread.Start();
            IsOpened.WaitOne();
            double Duration = mediaPlayer.currentMedia.duration;
            mediaPlayer.close();
            return (Duration);
        }
        private void Play()
        {
            mediaPlayer.OpenStateChange += MediaPlayer_OpenStateChange;
            mediaPlayer.URL = Url;
        }
        private void PlayInNewPlayer()
        {
            mediaPlayer = new WindowsMediaPlayer();
            mediaPlayer.OpenStateChange += MediaPlayer_OpenStateChange;
            mediaPlayer.URL = Url ;
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
        public void CheckUrlsAccessibility(Excel.Range RangeForCheck)//专用于检查乂学的视频链接有效性
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            Excel.Range RangeForReturn = RangeForCheck.Offset[0, RangeForCheck.Columns.Count];
            Url url = new Url();
            //string[,] Availability = new string[m, n];并不好使，网络出错会导致半途而废，很容易得不偿失，所以这里别用数组
            int sum = 0;
            int t = 0;
            for (int i = 1; i <= RangeForCheck.Rows.Count; i++)
            {
                if (i>8)
                {
                    CommonExcel.window.SmallScroll(1);//舒适地滚动
                    //System.Windows.Forms.Application.DoEvents();//不知道为啥，一DoEvents就卡死。不过等以后技术进步了，还是不要DoEvents了吧
                }
                for (int j = 1; j <= RangeForCheck.Columns.Count; j++)
                {
                    if (RangeForCheck[i, j].value != "无" && !string.IsNullOrWhiteSpace(RangeForCheck[i, j].value))
                    {
                        RangeForReturn[i, j].value = "正在验证有效性……";//实测降低25%性能，但是值得
                        sum++;
                        url.Value = RangeForCheck[i, j].value;
                        
                        Task<bool> checkAccessibilityTask =url.CheckAccessibilityAsync();
                        if (checkAccessibilityTask.Result)//用google试过，不翻墙会引发socket异常，暂时不需要处理
                        {
                            RangeForReturn[i, j].value = "有";
                        }
                        else
                        {
                            RangeForReturn[i, j].value = "无";
                            t++;
                        }
                    }
                };
            }
            //AvailabilityRange.Value = Availability;暂时不要了，找到好办法再说
            //WaitHandle.WaitAll();
            //stopwatch.Stop();
            //System.Windows.Forms.MessageBox.Show(@"耗时" + stopwatch.Elapsed.TotalSeconds + "秒" +
            //                                        "完成了" + sum + "个链接的有效性验证，其中" + t + "个无效");
        }
        public void CheckVideosLength(Excel.Range RangeForCheckVideoLength)
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            Excel.Range RangeForReturn = RangeForCheckVideoLength.Offset[0, RangeForCheckVideoLength.Columns.Count];
            AutoResetEvent[] Writers = new AutoResetEvent[5];
            int t = 0;
            for (int i = 1; i <= RangeForCheckVideoLength.Rows.Count; i++)
            {
                if (i > 8)
                {
                    CommonExcel.window.SmallScroll(1);//舒适地滚动
                }
                for (int j = 1; j <= RangeForCheckVideoLength.Columns.Count; j++)
                {
                    MyMediaPlayer myMediaPlayer = new MyMediaPlayer();
                    WindowsMediaPlayer mediaPlayer = new WindowsMediaPlayer();
                    RangeForReturn[i, j].value = "正在检测时长……";//实测轻微降低了性能，占比很小
                    RangeForReturn[i, j].value = myMediaPlayer.GetDuration(RangeForCheckVideoLength.Cells[i, j].value,ref mediaPlayer);
                    t++;
                }
            }
            //mediaPlayer.close();
            stopwatch.Stop();
            //System.Windows.Forms.MessageBox.Show(@"耗时" + stopwatch.Elapsed.TotalSeconds + "秒，" +
            //                                    "共选中了"+RangeForCheckVideoLength.Count+"个单元格，"+
            //                                    "成功完成了" + t + "个视频时长的检测");
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
