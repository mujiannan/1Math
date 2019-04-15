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
    public class CommonExcel
    {
        public delegate void DShedule(double Percent);
        public event DShedule SheduleChange;
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
        public string[,] ReadAsStr(Excel.Range FromRange)
        {
            int m = FromRange.Rows.Count;
            int n = FromRange.Columns.Count;
            string[,] Values = new string[m, n];
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    Values[i, j] = (string)FromRange[i + 1, j + 1].value;
                }
            }
            return (Values);
        }
        public void Write(object[,] Values,Excel.Range InRange)
        {
            CommonExcel.ExApp.ScreenUpdating = false;
            int m = InRange.Rows.Count;
            int n = InRange.Columns.Count;
            double Sum = InRange.Count;//这次没掉坑里
            int sum = 0;
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    sum++;
                    SheduleChange.Invoke(sum / Sum);
                    InRange[i+1,j+1].value=Values[i, j];
                }
            }
            ExApp.ScreenUpdating = true;
        }
        public object[,] ReadAntiMerge(Excel.Range FromRange)
        {
            int m = FromRange.Rows.Count;
            int n = FromRange.Columns.Count;
            object[,] Values = new object[m, n];
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    Values[i, j] = FromRange[i + 1, j + 1].MergeArea.Cells[1, 1].value;
                }
            }
            return (Values);
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
            CommonExcel CE = new CommonExcel();
            string[,] Urls = CE.ReadAsStr(CE.SelectedRange);
            SheduleChange.Invoke(0.1);
            bool[,] Accessibilities = new bool[CE.m, CE.n];
            Url url = new Url();
            int Sum = CE.SelectedRange.Count;
            int sum = 0;
            int t = 0;
            try
            {
                for (int i = 0; i < Urls.GetLength(0); i++)
                {
                    for (int j = 0; j < Urls.GetLength(1); j++)
                    {
                        sum++;
                        SheduleChange.Invoke(0.1 + 0.9 * sum / Sum);
                        url.Str = Urls[i, j];
                        if (url.Accessibility)
                        {
                            Accessibilities[i, j] = true;
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
                MessageChange.Invoke("测试完毕，回写Excel……");
                CE.WriteInRange.Value = Accessibilities;
                MessageChange.Invoke("回写完毕");
            }
            stopwatch.Stop();
            MessageChange.Invoke(@"耗时" + stopwatch.Elapsed.Seconds + "秒，"
                                                    + "完成了" + sum + "个链接的有效性验证，其中" + t + "个无效");
        }
        public void CheckVideosLength()
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            CommonExcel CE = new CommonExcel();
            string[,] Urls = CE.ReadAsStr(CE.SelectedRange);
            SheduleChange.Invoke(0.03);
            int Sum = Urls.Length;
            int sum = 0;
            int t = 0;
            MyMediaPlayer myMediaPlayer = new MyMediaPlayer();
            Url url = new Url();
            double[,] Durations = new double[CE.m, CE.n];
            for (int i = 0; i < Urls.GetLength(0); i++)
            {
                for (int j = 0; j < Urls.GetLength(1); j++)
                {
                    sum++;
                    url.Str = Urls[i, j];
                    Durations[i,j]=myMediaPlayer.GetDuration(url.Str);
                    MessageChange.Invoke(url.Str + "的时长为" + Durations[i, j] + "秒");
                    SheduleChange.Invoke(0.03+0.97*sum / Sum);
                    t++;
                }
            }
            MessageChange.Invoke("测试完毕，回写Excel");
            CE.SelectedRange.Offset[0, 2 * CE.n].Value = Durations;
            stopwatch.Stop();
            MessageChange.Invoke(@"耗时" + stopwatch.Elapsed.Seconds + "秒，" +
                                                "共选中了"+Sum+"个单元格，"+
                                                "成功完成了" + t + "个视频时长的检测");
        }
        public void AntiMerge()
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            MessageChange.Invoke("准备开始，从Excel中读取中……");//假的
            CommonExcel CE = new CommonExcel();
            
            CE.SheduleChange += CE_Shedule;
            object[,] RealValues = new object[CE.m,CE.n];
            SheduleChange.Invoke(0.01);
            MessageChange.Invoke("填充被合并的单元格中……");
            RealValues = CE.ReadAntiMerge(CE.SelectedRange);
            SheduleChange.Invoke(0.06);
            MessageChange.Invoke("取消合并单元格中……");
            CE.SelectedRange.UnMerge();
            SheduleChange.Invoke(0.08);
            MessageChange("逐个回写变体数据中，速度较慢，请耐心等候……");
            CE.Write(RealValues, CE.SelectedRange);
            stopwatch.Stop();
            MessageChange.Invoke("大功告成！共耗时"+stopwatch.Elapsed.Seconds+"秒");
            //CE.WriteInRange[2,1].Value = RealValues[1,0];
        }

        private void CE_Shedule(double Percent)
        {
            SheduleChange.Invoke(0.08+0.92*Percent);
        }

        //private object RealValue(Excel.Range range)
        //{
        //    if (range.MergeCells&&range.Value==null)
        //    {
        //        return (RealValue(range.Offset[-1, 0]));
        //    }
        //    else
        //    {
        //        return (range.Value);
        //    }
        //}
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
