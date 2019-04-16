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
    public class Tasks
    {

        private EventHandler Shutdown;
        private EventHandler Startup;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }
        public delegate void DelegateChangeStatus<T>(T item);
        public event DelegateChangeStatus<string> MessageChange;
        public event DelegateChangeStatus<double> ScheduleChange;
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
            ScheduleChange.Invoke(0.1);
            try
            {
                for (int i = 1; i <= Urls.GetLength(0); i++)
                {
                    for (int j = 1; j <= Urls.GetLength(1); j++)
                    {
                        sum++;
                        ScheduleChange.Invoke(0.1 + 0.9 * sum / Sum);
                        url.Str = Urls[i, j].ToString();
                        if (url.Accessibility)
                        {
                            Accessibilities[i - 1, j - 1] = true;
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
                application.Selection.OffSet[0, application.Selection.Columns.Count].Value = Accessibilities;
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
            double[,] Durations = new double[application.Selection.Rows.Count, application.Selection.Columns.Count];
            ScheduleChange.Invoke(0.03);
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
                        ScheduleChange.Invoke(0.03 + 0.97 * sum / Sum);
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
                                                "共选中了" + Sum + "个单元格，" +
                                                "成功完成了" + t + "个视频时长的检测");
        }
        public void AntiMerge()
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            MessageChange.Invoke("在选区中探寻合并单元格……");
            ArrayList MergedRanges = new ArrayList();
            Excel.Application application = Globals.ThisAddIn.Application;
            application.ScreenUpdating = false;
            Excel.Range RangeNeedUnMerge = application.Selection;
            if (RangeNeedUnMerge.Count == 1)
            {
                MessageChange.Invoke("只选中了一个单元格，默认处理整张工作表！");
                RangeNeedUnMerge = application.ActiveSheet.UsedRange;
            }
            application.FindFormat.MergeCells = true;
            Excel.Range Result = RangeNeedUnMerge.Find(What: "", After: RangeNeedUnMerge.Cells[1, 1], SearchFormat: true);
            Excel.Range FirstResult = Result;
            Excel.Range MergedRange = FirstResult;
            if (FirstResult == null)
            {
                MessageChange.Invoke("没有发现合并单元格！");
                ScheduleChange.Invoke(1);
                application.ScreenUpdating = true;
                return;//直接return舒服一些，别else了……
            }
            else
            {
                //卧槽，还是要else，太坑了。随便查找一下，如果只选中一块合并单元格，竟然是能超出当前选取的……
                //excel这种设定，倒也合理，直接把整个合并单元格区域当成起点，跑下一段去了。但是，它跟vba不一致啊……
                if (FirstResult.Row > (RangeNeedUnMerge.Row + RangeNeedUnMerge.Rows.Count - 1))
                {
                    MergedRanges.Add(RangeNeedUnMerge);
                    goto Fool;
                }
            }
            MergedRange = FirstResult.MergeArea;
            int t = 0;
            ScheduleChange.Invoke(0.1);
            do
            {
                t++;
                ScheduleChange.Invoke(0.5 - 0.4 / t);
                MergedRanges.Add(MergedRange);
                Result = RangeNeedUnMerge.Find(What: "", After: Result, SearchFormat: true); ;//这里的接龙很巧妙，但也很坑
                MergedRange = Result.MergeArea;
            } while (MergedRange != null && MergedRange.Cells[1, 1].Address != FirstResult.Address);//卧槽，必须用do while，如果先判断，肯定就跳过去了。我写的逻辑太坑，一层套一层，每一步都对后续步骤有深渊影响
        Fool:
            MessageChange.Invoke("已找到所有合并单元格，正在安全拆分……");
            ScheduleChange.Invoke(0.5);
            RangeNeedUnMerge.UnMerge();
            ScheduleChange.Invoke(0.55);
            int Sum = MergedRanges.Count;
            int sum = 0;
            foreach (Excel.Range range in MergedRanges)
            {
                ScheduleChange.Invoke(0.55 + 0.4 * ((++sum)/ Sum));
                range.Value = range.Cells[1, 1];
            }
            application.ScreenUpdating = true;
            ScheduleChange.Invoke(1);
            stopwatch.Stop();
            MessageChange.Invoke("耗时" + (stopwatch.Elapsed.TotalSeconds).ToString() + "秒");
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
    
}
