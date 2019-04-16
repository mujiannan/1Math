using System.Net;
using System.Net.Http;
using System;
using System.Threading.Tasks;
using WMPLib;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections;
namespace _1Math
{
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
            checkTask =Check();
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
    public class Tasks
    {
        
        private EventHandler Shutdown;
        private EventHandler Startup;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }
        public delegate void DelegateChangeStatus<T>(T item);
        public event DelegateChangeStatus<string> MessageChange;
        public event DelegateChangeStatus<double> SheduleChange;
        //private double TaskShedule;

        public void CheckUrlsAccessibility()//专用于检查乂学的视频链接有效性
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            MessageChange.Invoke("准备资源中……");
            Excel.Application application = Globals.ThisAddIn.Application;
            object[,] Urls = application.Selection.Value;
            bool[,] Accessibilities = new bool[application.Selection.Rows.Count, application.Selection.Columns.Count];
            Url url = new Url();
            int Sum = application.Selection.Count;
            int sum = 0;
            int t = 0;
            SheduleChange.Invoke(0.1);
            try
            {
                for (int i = 1; i <=Urls.GetLength(0); i++)
                {
                    for (int j = 1; j <= Urls.GetLength(1); j++)
                    {
                        sum++;
                        SheduleChange.Invoke(0.1 + 0.9 * sum / Sum);
                        url.Str = Urls[i, j].ToString();
                        if (url.Accessibility)
                        {
                            Accessibilities[i-1, j-1] = true;
                            MessageChange.Invoke(url.Str + "成功");
                        }
                        else
                        {
                            MessageChange.Invoke(url.Str + "失败");
                            t++;
                        }
                    }
                }
            }
            finally
            {
                application.Selection.OffSet[0,application.Selection.Columns.Count].Value = Accessibilities;
            }
            stopwatch.Stop();
            MessageChange.Invoke(@"耗时" + stopwatch.Elapsed.Seconds + "秒，"
                                                    + "完成了" + sum + "个链接的有效性验证，其中" + t + "个无效");
        }
        public void CheckVideosLength()
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            MessageChange.Invoke("准备资源中……");
            Excel.Application application = Globals.ThisAddIn.Application;
            object[,] Urls = application.Selection.Value;
            
            int Sum = Urls.Length;
            int sum = 0;
            int t = 0;
            MyMediaPlayer myMediaPlayer = new MyMediaPlayer();
            Url url = new Url();
            double[,] Durations = new double[application.Selection.Rows.Count,application.Selection.Columns.Count];
            SheduleChange.Invoke(0.03);
            try
            {
                for (int i = 1; i <= Urls.GetLength(0); i++)
                {
                    for (int j = 1; j <= Urls.GetLength(1); j++)
                    {
                        sum++;
                        url.Str = Urls[i, j].ToString();
                        Durations[i - 1, j - 1] = url.Accessibility ? myMediaPlayer.GetDuration(url.Str) : 0;
                        MessageChange.Invoke(url.Str + "的时长为" + Durations[i - 1, j - 1] + "秒");
                        SheduleChange.Invoke(0.03 + 0.97 * sum / Sum);
                        t++;
                    }
                }
            }
            finally
            {
                application.Selection.Offset[0, 2 * application.Selection.Columns.Count].Value = Durations;
            }
            stopwatch.Stop();
            MessageChange.Invoke(@"耗时" + stopwatch.Elapsed.Seconds + "秒，" +
                                                "共选中了"+Sum+"个单元格，"+
                                                "成功完成了" + t + "个视频时长的检测");
        }
        private object[,] ReadAntiMerge(Excel.Range FromRange)
        {
            int m = FromRange.Rows.Count;
            int n = FromRange.Columns.Count;
            object[,] Values = new object[m, n];
            double Sum = FromRange.Count;
            int sum = 0;
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    sum++;
                    Values[i, j] = FromRange[i+1,j+1].MergeCells ? (object)FromRange[i + 1, j + 1].MergeArea.Cells[1, 1].value : (object)FromRange[i + 1, j + 1].value;
                    //SheduleChange.Invoke(0.05+0.75*sum / Sum);
                }
            }
            return (Values);
        }
        public void AntiMerge()
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            Excel.Application application = Globals.ThisAddIn.Application;
            application.ScreenUpdating = false;
            double Sum=application.Selection.Count;
            int sum = 0;
            foreach (Excel.Range range in application.Selection)
            {
                sum++;
                if (range.MergeCells)
                {
                    Excel.Range MergedRange = range.MergeArea;
                    MergedRange.UnMerge();
                    MergedRange.Value = MergedRange.Cells[1, 1].value;
                }
                SheduleChange.Invoke(sum / Sum);
            }
            application.ScreenUpdating = true;
            stopwatch.Stop();
            MessageChange.Invoke("耗时"+(stopwatch.Elapsed.TotalSeconds).ToString()+"秒");
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
