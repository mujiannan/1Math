using System.Net;
using System.Net.Http;
using System;
using System.Threading.Tasks;
using WMPLib;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections;
using System.Windows.Media;
namespace _1Math
{
    public class NetTasks
    {
        protected System.Diagnostics.Stopwatch stopwatch;
        protected Excel.Application application;
        protected Excel.Range rangeForReturn;
        private int ThreadsCount;
        protected int threadsLimit;
        private EventHandler Shutdown;
        private EventHandler Startup;
        public delegate void DelegateChangeStatus<T>(T item);
        public event DelegateChangeStatus<string> MessageChange;
        public event DelegateChangeStatus<double> ProgressChange;
        private delegate void DTestUrl(Url url, int i, int j);
        protected int m, n;
        private int x = 0, y = 1;
        volatile DTestUrl[,] dTestUrls;
        object[,] UrlsRange;
        protected double Sum;//总任务量
        private int sum = 0;//完成任务量
        public NetTasks()
        {
            stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            application = Globals.ThisAddIn.Application;
            UrlsRange = application.Selection.Value;
            Sum = UrlsRange.Length;
            m = UrlsRange.GetLength(0);
            n = UrlsRange.GetLength(1);
            dTestUrls = new DTestUrl[m, n];
        }
        public void Start()
        {
            
            for (int i = 0; i < threadsLimit; i++)
            {
                Next();
            }
        }
        protected virtual void Complete()//结束，必须覆写
        {

        }
        public void Report(string Message)
        {
            MessageChange.BeginInvoke(Message,null,null);
        }
        protected void CompleteOne()
        {
            sum++;
            ProgressChange.BeginInvoke(sum / Sum, null, null);
            ThreadsCount--;
            if (sum < Sum)
            {
                Next();
            }
            else
            {
                ProgressChange.BeginInvoke(1, null, null);
                stopwatch.Stop();
                Complete();
            }

        }
        private void Next()//封闭着就行了，完全不用动
        {
            if (x < m)
            {
                x++;
            }
            else if (y < n)
            {
                y++;
            }
            else
            {
                return;
            }
            ThreadsCount++;
            Url url = new Url
            {
                Str = UrlsRange[x, y].ToString()
            };
            dTestUrls[x - 1, y - 1] = new DTestUrl(Work);
            dTestUrls[x - 1, y - 1].BeginInvoke(url, x, y, null, null);
        }

        protected virtual void Work(Url url, int i, int j)//工作方法，必须覆盖，不然就直接结束
        {
            CompleteOne();
        }
    }
    public class Accessibility : NetTasks
    {
        private bool[,] Results;
        public Accessibility()
        {
            threadsLimit = 4;
            Results  = new bool[m, n];
            rangeForReturn = application.Selection.OffSet[0, n];
        }
        private int InAccessibleUrlsCount;
        protected override void Complete()
        {
            rangeForReturn.Value = Results;
            Report("耗时" + stopwatch.Elapsed.TotalSeconds.ToString() + "秒，共验证了" + Sum + "个视频的有效性，其中" + InAccessibleUrlsCount + "个无效");
        }
        protected override void Work(Url url, int i, int j)
        {
            bool accessibility;
            accessibility = url.Accessibility;
            Results[i - 1, j - 1] = accessibility;
            if (!accessibility)
            {
                InAccessibleUrlsCount++;
            }
            //MessageChange.Invoke(url.Str + "验证结果：" + accessibility);
            CompleteOne();
        }
    }
    public class VideoLength : NetTasks
    {
        private volatile double[,] Results;
        public VideoLength()
        {
            threadsLimit = 2;
            Results = new double[m, n];
            rangeForReturn = application.Selection.OffSet[0, 2 * n];
        }
        private volatile int Success;
        protected override void Complete()
        {
            rangeForReturn.Value = Results;
            Report("耗时" + stopwatch.Elapsed.TotalSeconds.ToString() + "秒，成功获取了" + Success + "个视频的时长");
        }
        protected override void Work(Url url, int i, int j)
        {
            Func<double> func = (() =>
            {
                MediaPlayer mediaPlayer = new MediaPlayer();
                mediaPlayer.Open(new Uri(url.Str));
                DateTime start = DateTime.Now;
                TimeSpan timeSpan;
                do
                {
                    Thread.Sleep(100);
                    timeSpan = DateTime.Now - start;
                } while (!mediaPlayer.NaturalDuration.HasTimeSpan&&timeSpan.TotalSeconds<2);
                double Duration = mediaPlayer.NaturalDuration.HasTimeSpan ? mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds : 0;
                mediaPlayer.Close();
                if (Duration>0)
                {
                    Success++;
                } 
                return (Duration);
            });
            Results[i - 1, j - 1] = func();
            CompleteOne();
        }
    }

    public class Tasks
    {
        Excel.Application application;
        Excel.Range rangeForReturn;
        private int ThreadsCount;
        int threadsLimit;
        private EventHandler Shutdown;
        private EventHandler Startup;
        public delegate void DelegateChangeStatus<T>(T item);
        public event DelegateChangeStatus<string> MessageChange;
        public event DelegateChangeStatus<double> ProgressChange;
        private delegate void DTestUrl(Url url,int i,int j);
        private int m, n;
        private int x=0, y=1;
        bool[,] accessibilities;
        DTestUrl[,] dTestUrls;
        object[,] UrlsRange;
        private double Sum;
        private int sum=0;
        private int InAccessibleUrlsCount;
        private delegate void DComplete();
        private event DComplete CompleteOne;
        private event DComplete Complete;
        private delegate void DCompleteOne(int sender);
        private System.Diagnostics.Stopwatch stopwatch;
        public enum TaskName
        {
            CheckUrlsAccessibility,CheckVideosLength,AntiMerge
        }
        public Tasks(TaskName taskName)
        {
            stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            Initialize(taskName);
        }
        private void Initialize(TaskName taskName)
        {
            application = Globals.ThisAddIn.Application;

            
            if ((int)taskName<2)
            {
                UrlsRange = application.Selection.Value;
                Sum = UrlsRange.Length;
                m = UrlsRange.GetLength(0);
                n = UrlsRange.GetLength(1);
                rangeForReturn = application.Selection.OffSet[0, n];
                dTestUrls = new DTestUrl[m, n];
            }
        }
        public void CheckUrlsAccessibility()
        {
            accessibilities = new bool[m, n];
            CompleteOne += Tasks_CompleteOneAccessibility;
            Complete += Tasks_CompleteAccessibilities;
            threadsLimit = 4;
            for (int i = 0; i < threadsLimit; i++)
            {
                NextAccessibility();
            }
        }
        private void Tasks_CompleteAccessibilities()
        {
            rangeForReturn.Value= accessibilities;
            ProgressChange.BeginInvoke(1,null,null);
            stopwatch.Stop();
            MessageChange.BeginInvoke("耗时" + stopwatch.Elapsed.TotalSeconds.ToString() + "秒，共验证了" + Sum + "个视频的有效性，其中" + InAccessibleUrlsCount + "个无效",null,null);
        }
        private void Tasks_CompleteOneAccessibility()
        {
            sum++;
            ProgressChange.BeginInvoke(sum / Sum, null, null);
            ThreadsCount--;
            if (sum<Sum)
            {
                NextAccessibility();
            }
            else
            {
                Complete.BeginInvoke(null,null);
            }

        }
        private void NextAccessibility()
        {
            if (x < m)
            {
                x++;
            }
            else if (y < n)
            {
                y++;
            }
            else
            {
                return;
            }
            ThreadsCount++;
            Url url = new Url
            {
                Str = UrlsRange[x, y].ToString()
            };
            dTestUrls[x - 1, y - 1] = new DTestUrl(WriteAccessibilityIn);
            dTestUrls[x - 1, y - 1].BeginInvoke(url, x, y, null, null);
        }
        private void WriteAccessibilityIn(Url url,int i,int j)
        {
            bool accessibility;
            accessibility = url.Accessibility;
            accessibilities[i - 1, j - 1] = accessibility;
            if (!accessibility)
            {
                InAccessibleUrlsCount++;
            }
            //MessageChange.Invoke(url.Str + "验证结果：" + accessibility);
            CompleteOne.Invoke();
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
            ProgressChange.BeginInvoke(0.03,null,null);
            try
            {
                for (int i = 1; i <= Urls.GetLength(0); i++)
                {
                    for (int j = 1; j <= Urls.GetLength(1); j++)
                    {
                        sum++;
                        url.Str = Urls[i, j].ToString();
                        int RetryTimes = 0;
                        Retry:
                        try
                        {
                            Durations[i - 1, j - 1] = myMediaPlayer.GetDuration(url.Str);
                        }
                        catch (Exception)
                        {
                            if (RetryTimes<5)
                            {
                                RetryTimes++;
                                goto Retry;
                            }
                        }
                        MessageChange.Invoke(url.Str + "的时长为" + Durations[i - 1, j - 1] + "秒");
                        ProgressChange.BeginInvoke(0.03 + 0.97 * sum / Sum,null,null);
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
            MessageChange.Invoke("在选区中探寻合并单元格……");
            application.ScreenUpdating = false;
            ProgressChange.Invoke(0.05);
            MergedAreas mergedAreas = new MergedAreas();
            ProgressChange.Invoke(0.5);
            MessageChange.Invoke("已找到所有合并单元格，正在安全拆分……");
            mergedAreas.SafelyUnMergeAndFill();
            application.ScreenUpdating = true;
            ProgressChange.Invoke(1);
            stopwatch.Stop();
            MessageChange.Invoke("耗时" + (stopwatch.Elapsed.TotalSeconds).ToString() + "秒");
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
    class MergedAreas//实例化后，直接运行SafelyUnMergedAndFill方法即可。拆分的目标默认为当前Selection。目标区域为Single时，则自动将目标区域更改为整个活动工作表
    {
        private Excel.Application application;
        private Excel.Range Target;
        private ArrayList mergedAreas;

        public MergedAreas()
        {
            application = Globals.ThisAddIn.Application;
            Target = application.Selection;
            GetAsArrayList();
        }
        public MergedAreas(Excel.Range In)//此重载提供了将这个类用于快速获取合并单元格区域的可能性
        {
            application = In.Application;
            Target = In;
            GetAsArrayList();
        }
        public ArrayList ToArrayList()//类本身当然不能作为数组，但我可以为其添加ToArrayList方法，伪装一下
        {
            return mergedAreas;
        }
        public void SafelyUnMergeAndFill()
        {
            Target.UnMerge();
            foreach (Excel.Range range in mergedAreas)
            {
                range.Value = range.Cells[1, 1];
            }
        }
        private void GetAsArrayList()
        {
            mergedAreas = new ArrayList();
            if (Target.Count == 1)
            {
                //只选择了一个单元格，自动将搜寻区域拓展至其所在的整张工作表
                Target = Target.Worksheet.UsedRange;//这样的设定会使我们开发出更便于使用的VSTO
            }
            application.FindFormat.MergeCells = true;
            Excel.Range Result = Target.Find(What: "", After: Target.Cells[1, 1], SearchFormat: true);
            Excel.Range FirstResult = Result;
            Excel.Range MergedArea = Result;
            if (FirstResult == null)
            {
                //没有发现合并单元格！
                mergedAreas = null;
                return;
            }
            else
            {
                //卧槽，还是要else，太坑了。随便查找一下，如果只选中一块合并单元格，竟然是会跳出当前选区的……
                //excel这种设定，倒也合理，直接把整个合并单元格区域当成起点，跑下一段去了。但是，它跟vba不一致啊……
                if (FirstResult.Row > (Target.Row + Target.Rows.Count - 1))
                {
                    mergedAreas.Add(Target);
                    return;
                }
            }
            MergedArea = FirstResult.MergeArea;
            int t = 0;
            do
            {
                t++;
                mergedAreas.Add(MergedArea);
                Result = Target.Find(What: "", After: Result, SearchFormat: true); ;//这里的接龙很巧妙，但也很坑。我还尝试着用FindNext，但是出现了一点问题。
                MergedArea = Result.MergeArea;
            } while (MergedArea != null && MergedArea.Cells[1, 1].Address != FirstResult.Address);
        }
    }
    public class Url
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
            checkClient.Timeout = new TimeSpan(0, 0, 1);
            try
            {
                HttpResponseMessage httpResponseMessage = await checkClient.GetAsync(str, HttpCompletionOption.ResponseHeadersRead);
                accessibility = httpResponseMessage.IsSuccessStatusCode;
            }
            catch (Exception)
            {
                accessibility = false;
            }
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
    class MyMediaPlayer:IDisposable
    {
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool Disposing)
        {
            if (Disposing)
            {
                IsOpened.Dispose();
                mediaPlayer.close();
                mediaPlayer = null;
            }
            else
            {
                mediaPlayer.close();
            }
        }
        ~MyMediaPlayer()
        {
            Dispose(false);
        }
        WindowsMediaPlayer mediaPlayer;
        private string mediaUrl;
        private AutoResetEvent IsOpened;
        public MyMediaPlayer()
        {
            IsOpened = new AutoResetEvent(false);
            Retry:
            try
            {
                mediaPlayer = new WindowsMediaPlayer();
                mediaPlayer.OpenStateChange += MediaPlayer_OpenStateChange;
            }
            catch (Exception)
            {
                Thread.Sleep(500);
                goto Retry;
            }
            
            
        }
        public double GetDuration(string url)
        {
            Thread PlayThread = new Thread(Play);
            try
            {
                mediaUrl = url;
                PlayThread.Start();
                IsOpened.WaitOne(2000);
                double Duration = mediaPlayer.currentMedia.duration;
                return Duration;
            }
            catch (Exception)
            {
                return 0;
            }

        }
        private void Play()
        {
            try
            {
                mediaPlayer.URL = mediaUrl;
            }
            catch (Exception)
            {
                throw;
            }
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
