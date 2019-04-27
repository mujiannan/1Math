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
    
    public static class CE
    {
        private static System.Diagnostics.Stopwatch stopwatch;
        public static string Elapse
        {
            get
            {
                return (stopwatch.Elapsed.TotalSeconds.ToString());
            }
        }
        public static Excel.Application App = Globals.ThisAddIn.Application;
        public static Excel.Range Selection
        {
            get
            {
                return (App.Selection);
            }
        }
        public static void StartTask()
        {
            stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            App.ScreenUpdating = false;
        }
        public static void EndTask()
        {
            App.ScreenUpdating = true;
            stopwatch.Stop();
        }
    }
    public abstract class NetTask
    {
        protected Thread[] threads;
        protected Excel.Range rangeForReturn;
        protected int threadsLimit;
        public delegate void DelegateChangeStatus<T>(T item);
        public event DelegateChangeStatus<string> MessageChange;
        public event DelegateChangeStatus<double> ProgressChange;
        protected int m, n;
        private volatile int x = 0, y = 1;
        protected object[,] UrlsRange;//不需要锁
        protected double Sum;//总任务量
        private volatile int sum = 0;//完成任务量
        private readonly int[] HasNoNext =new int[2]{0,0};
        protected NetTask()
        {
            CE.StartTask();
            if (CE.Selection.Count > 1)
            {
                UrlsRange=CE.Selection.Value;
            }
            else
            {
                UrlsRange = (object[,])Array.CreateInstance(typeof(object), new int[2] { 1, 1 }, new int[2] { 1, 1 });
                UrlsRange[1, 1] =CE.Selection.Cells[1,1].Value;
                
            }
            Sum = UrlsRange.Length;
            m = UrlsRange.GetLength(0);
            n = UrlsRange.GetLength(1);
        }
        public void Start()
        {
            for (int i = 0; i < threads.Length; i++)
            {
                threads[i] = new Thread(Work);
                threads[i].Start();
            };
        }
        private void Finish()
        {
            ProgressChange(1);
            CE.EndTask();
        }
        protected virtual void Complete() { }//结束，必须覆写
        protected void Report(string Message)
        {
            MessageChange.BeginInvoke(Message,null,null);
        }
        protected void CompleteOne()
        {
            sum++;
            if (sum < Sum)
            {
                ProgressChange.BeginInvoke(sum / Sum, null, null);
            }
            else
            {
                Finish();
                Complete();
            }

        }
        protected int[] GetNext()//封闭着就行了，完全不用动
        {
            if (x < m)
            {
                x++;
            }
            else if (y < n)
            {
                x = 1;
                y++;
            }
            else
            {
                return(HasNoNext);
            }
            return (new int[2] { x, y });
        }
        protected virtual void Work() => CompleteOne();//工作方法，必须覆盖，不然就直接结束
    }
    public class Accessibility : NetTask
    {
        private bool[,] results;//这个也许要锁，不太确定
        public Accessibility()
        {
            threadsLimit = 10;
            threads = new Thread[threadsLimit];
            results  = new bool[m, n];
            rangeForReturn = CE.Selection.Offset[0, n];
        }
        private int inAccessibleUrlsCount;
        protected override void Complete()
        {
            rangeForReturn.Value = results;
            Report("耗时" + CE.Elapse + "秒，共验证了" + Sum + "个视频的有效性，其中" + inAccessibleUrlsCount + "个无效");
        }
        protected override void Work()
        {
            int[] next = GetNext();
            bool accessibility;
            int i, j;
            Url url = new Url();
            while (next[0] != 0)
            {
                i = next[0];
                j = next[1];
                url.SetReferTo(UrlsRange[i, j].ToString());
                accessibility = url.Accessibility;
                results[i - 1, j - 1] = accessibility;
                if (!accessibility)
                {
                    inAccessibleUrlsCount++;
                }
                CompleteOne();
                next = GetNext();
            };
        }
    }
    public class VideoLength : NetTask
    {
        private double[,] results;
        public VideoLength()
        {
            threadsLimit = 1;
            threads = new Thread[threadsLimit];
            results = new double[m, n];
            rangeForReturn = CE.Selection.Offset[0,2*n];
        }
        private int success;
        protected override void Complete()
        {
            rangeForReturn.Value = results;
            Report("耗时" + CE.Elapse + "秒，测试了" + Sum + "个视频的时长，其中" + success + "个测试成功");
        }
        protected override void Work()
        {
            int[] next = GetNext();
            double duration;
            int i, j;
            MyMediaPlayer mediaPlayer = new MyMediaPlayer();
            while (next[0] != 0)
            {
                i = next[0];
                j = next[1];
                duration = mediaPlayer.GetDuration(UrlsRange[i, j].ToString());
                results[i - 1, j - 1] = duration;
                if (duration>0)
                {
                    success++;
                }
                CompleteOne();
                next = GetNext();
            };
            mediaPlayer.Dispose();
        }
    }

    public class Tasks
    {
        public delegate void DelegateChangeStatus<T>(T item);
        public event DelegateChangeStatus<string> MessageChange;
        public event DelegateChangeStatus<double> ProgressChange;

        public void AntiMerge()
        {
            CE.StartTask();
            MessageChange.Invoke("在选区中探寻合并单元格……");
            ProgressChange.Invoke(0.05);
            MergedAreas mergedAreas = new MergedAreas();
            ProgressChange.Invoke(0.5);
            MessageChange.Invoke("已找到所有合并单元格，正在安全拆分……");
            mergedAreas.SafelyUnMergeAndFill();
            ProgressChange.Invoke(1);
            CE.EndTask();
            MessageChange.Invoke("耗时" + CE.Elapse + "秒");
        }
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
        private string str;
        public void SetReferTo(string value)
        {
            if (value != str)
            {
                checkStatus = CheckStatus.Null;//这样可以不需要反复实例化就可以验证有效性，不知道会不会节省资源
                str = value;
            }
        }
        public new string ToString()//假的
        {
            return str;
        }
        public Url(string url)
        {
            str = url;
        }
        public Url()
        {
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
            Retry:
            try
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
            catch (Exception)
            {

                goto Retry;
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
                Thread.Sleep(200);
                goto Retry;
            }
            
            
        }
        public double GetDuration(string url)
        {
            Thread PlayThread = new Thread(Play);
            int RetryTimes = 0;
            Retry:
            try
            {
                mediaUrl = url;
                PlayThread.Start();
                IsOpened.WaitOne(5000);
                double Duration = mediaPlayer.currentMedia.duration;
                return Duration;
            }
            catch (Exception)
            {
                if (RetryTimes<3)
                {
                    RetryTimes++;
                    goto Retry;
                }
                else
                {
                    return 0;
                }
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
                return;
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
