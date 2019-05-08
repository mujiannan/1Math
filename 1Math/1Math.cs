using System.Net;
using System.Net.Http;
using System;
using System.Threading.Tasks;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections;
using System.Windows.Media;
namespace _1Math
{
    public delegate void CancelableMethod(CancellationToken token);
    public class BackGroundTask
    {
        private Task _backGroundTask;
        CancellationTokenSource _CTS = new CancellationTokenSource();
        public void Start(CancelableMethod cancelableMethod)
        {
            _backGroundTask = new Task(() =>
            {
                cancelableMethod(_CTS.Token);
            }, _CTS.Token);
            _backGroundTask.Start();
        }
        public BackGroundTask(IHasStatusReporter ObjHasStatusReporter)
        {
            StatusForm statusForm = new StatusForm();
            statusForm.Show();
            ObjHasStatusReporter.MessageChange += statusForm.MessageLabel_TextChange;
            ObjHasStatusReporter.ProgressChange += statusForm.ProgressBar_ValueChange;
            statusForm.FormClosing += StatusForm_FormClosing;
        }

        private void StatusForm_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            if (!_CTS.IsCancellationRequested)
            {
                _CTS.Cancel();
            }
            CE.EndTask();
            GC.Collect();
        }
    }
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
    public abstract class NetTask:IHasStatusReporter//此抽象类是针对excel中长耗时任务的一个多线程模板
    {
        protected CancellationToken Canceling;
        protected Thread[] threads;
        protected int threadsLimit=2;//默认双线程
        protected Excel.Range rangeForReturn;
        public event ChangeMessage MessageChange;
        public event ChangeProgress ProgressChange;
        protected int m, n;
        private volatile int x = 0, y = 1;
        protected object[,] UrlsRange;//不需要锁
        protected int Sum;//总任务量
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
        public void Start(CancellationToken cancellationToken=new CancellationToken())
        {
            Canceling = cancellationToken;
            Report("正在准备资源……");
            threadsLimit = System.Math.Min(threadsLimit, Sum);//这样可省事儿多了，根据任务数量与预设的线程上限共同确定线程数
            threads = new Thread[threadsLimit];
            for (int i = 0; i < threads.Length; i++)
            {
                threads[i] = new Thread(Work);
                threads[i].Start();
            };
            Report($"正在处理,线程数：{threadsLimit}个");
        }
        private void End()
        {
            Report(1);
            CE.EndTask();
        }//结束
        protected abstract void Complete();//完成全部任务
        protected void Report(string Message)
        {
            ChangeMessage changeMessage = MessageChange;
            if (changeMessage != null)
            {
                MessageChange(this, new MessageEventArgs(Message));
            }

        }
        protected void Report(double Progress)
        {
            ChangeProgress changeProgress = ProgressChange;
            if (changeProgress != null)
            {
                ProgressChange(this, new ProgressEventArgs(Progress));
            }
        }
        protected void CompleteOne()
        {
            sum++;
            if (sum < Sum)
            {
                Report(sum /(double)Sum);
            }
            else
            {
                Complete();
                End();
            }

        }
        protected int[] GetNext()//封闭着就行了，完全不用动
        {
            if (Canceling.IsCancellationRequested)//这里，巧妙地在工作线程每次领取下一个任务时检查任务是否被取消
            {
                return (HasNoNext);
            }
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
        protected abstract void Work();//工作方法
    }
    public class Accessibility : NetTask
    {
        private bool[,] results;//这个也许要锁，不太确定
        public Accessibility()
        {
            threadsLimit = Environment.ProcessorCount*2;
            results  = new bool[m, n];
            rangeForReturn = CE.Selection.Offset[0, n];
        }
        private int inAccessibleUrlsCount=0;
        protected override void Complete()
        {
            rangeForReturn.Value = results;
            Report($"耗时{CE.Elapse}秒，共验证了{Sum}个视频的有效性，其中{inAccessibleUrlsCount}个无效");
        }
        protected override void Work()
        {
            int[] next = GetNext();
            Url url = new Url();
            while (next[0] != 0)
            {
                int i = next[0];
                int j = next[1];
                url.SetReferTo(UrlsRange[i, j].ToString());
                bool accessibility = url.Accessibility;
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
            threadsLimit = Environment.ProcessorCount;
            results = new double[m, n];
            rangeForReturn = CE.Selection.Offset[0,2*n];
        }
        private volatile int success=0;
        protected override void Complete()
        {
            rangeForReturn.Value = results;
            Report($"耗时{CE.Elapse}秒，测试了{Sum}个视频的时长，其中{success}个测试成功");
        }
        protected override void Work()
        {
            int[] next = GetNext();
            DotNetPlayer dotNetPlayer = new DotNetPlayer();
            while (next[0] != 0)
            {
                int i = next[0];
                int j = next[1];
                double duration = dotNetPlayer.GetDuration(new Uri(UrlsRange[i, j].ToString()));
                if (duration>0)
                {
                    results[i - 1, j - 1] = duration;
                    success++;
                }
                CompleteOne();
                next = GetNext();
            };
            dotNetPlayer.Dispose();
        }
    }
    public class DotNetPlayer:IDisposable
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~DotNetPlayer()
        {
            Dispose(false);
        }
        private void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //呵呵，暂时还没有啥托管资源需要释放的
                }
                mediaPlayer.Close();
                mediaPlayer = null;
                disposed = true;
            }
            
        }
        MediaPlayer mediaPlayer;
        const double timeOut = 5;
        public DotNetPlayer()
        {
            mediaPlayer = new MediaPlayer();
        }
        public double GetDuration(Uri uri)
        {
            mediaPlayer.Open(uri);
            double duration = 0;
            DateTime start = DateTime.Now;
            TimeSpan timeSpan = new TimeSpan(0);
            do
            {
                Thread.Sleep(50);
                if (mediaPlayer.NaturalDuration.HasTimeSpan)
                {
                    duration = mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                    mediaPlayer.Stop();
                }
                timeSpan = DateTime.Now - start;
            } while (duration==0&&timeSpan.TotalSeconds<timeOut);
            return (duration);
        }
    }

    
    class MergeAreas:IHasStatusReporter//实例化后，直接运行SafelyUnMergedAndFill方法即可。拆分的目标默认为当前Selection。目标区域为Single时，则自动将目标区域更改为整个活动工作表
    {
        CancellationToken _cancellationToken;
        private readonly Excel.Application _application;
        private Excel.Range _target;
        private ArrayList _mergedAreas;
        public event ChangeMessage MessageChange;
        public event ChangeProgress ProgressChange;
        private ArrayList MergedAreas
        {
            get
            {
                if (_mergedAreas == null)
                {
                    GetMergedAreas();
                }
                return _mergedAreas;
            }
        }
        public MergeAreas()
        {
            CE.StartTask();
            _application = Globals.ThisAddIn.Application;
            _target = _application.Selection;
        }
        public MergeAreas(Excel.Range In)//此重载提供了将这个类用于快速获取合并单元格区域的可能性
        {
            CE.StartTask();
            _application = In.Application;
            _target = In;
        }
        public ArrayList ToArrayList()//类本身当然不能作为数组，但我可以为其添加ToArrayList方法，伪装一下
        {
            if (_mergedAreas == null)
            {
                GetMergedAreas();
            }
            ProgressChange(this,new ProgressEventArgs(1));
            return _mergedAreas;
        }
        public async void SafelyUnMergeAndFill(CancellationToken cancellationToken=new CancellationToken())
        {
            _cancellationToken = cancellationToken;
            if (_mergedAreas==null)
            {
                await Task.Run(new Action(GetMergedAreas));
                if (_mergedAreas==null)
                {
                    CE.EndTask();
                    MessageChange(this,new MessageEventArgs("找不到合并的单元格"));
                    ProgressChange(this,new ProgressEventArgs(1));
                    return;
                }
            }
            ProgressChange(this,new ProgressEventArgs(0.5));
            if (_cancellationToken.IsCancellationRequested)
            {
                return;
            }
            _target.UnMerge();
            int t = 0;
            double Sum = _mergedAreas.Count;
            foreach (Excel.Range range in _mergedAreas)
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    return;
                }
                t++;
                MessageChange(this, new MessageEventArgs($"取消合并中，第{t}个……"));
                ProgressChange(this, new ProgressEventArgs(0.5+t/Sum/2));
                range.Value = range.Cells[1, 1];//为什么这也能迭代……啥原因呢，不是说foreach不能这么来么
            }
            ProgressChange(this,new ProgressEventArgs(1));
            CE.EndTask();
            MessageChange(this,new MessageEventArgs($"大功告成，耗时{CE.Elapse}秒！"));
        }
        private void GetMergedAreas()
        {
            
            _mergedAreas = new ArrayList();
            if (_target.Count == 1)
            {
                MessageChange(this,new MessageEventArgs("只选择了一个单元格，自动将搜寻区域拓展至其所在的整张工作表"));
                _target = _target.Worksheet.UsedRange;//这样的设定会使我们开发出更便于使用的VSTO
            }
            _application.FindFormat.MergeCells = true;
            Excel.Range Result = _target.Find(What: "", After: _target.Cells[1, 1], SearchFormat: true);
            Excel.Range FirstResult = Result;
            Excel.Range MergedArea = Result;
            if (FirstResult == null)
            {
                MessageChange(this,new MessageEventArgs("没有发现合并单元格！"));
                _mergedAreas = null;
                return;
            }
            else
            {
                //卧槽，还是要else，太坑了。随便查找一下，如果只选中一块合并单元格，竟然是会跳出当前选区的……
                //excel这种设定，倒也合理，直接把整个合并单元格区域当成起点，跑下一段去了。但是，它跟vba不一致啊……
                if (FirstResult.Row > (_target.Row + _target.Rows.Count - 1))
                {
                    _mergedAreas.Add(_target);
                    return;
                }
            }
            MergedArea = FirstResult.MergeArea;
            int t = 0;
            do
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    return;
                }
                t++;
                MessageChange(this, new MessageEventArgs($"搜寻中，找到{t}处合并区域"));
                _mergedAreas.Add(MergedArea);
                Result = _target.Find(What: "", After: Result, SearchFormat: true); ;//这里的接龙很巧妙，但也很坑。我还尝试着用FindNext，但是出现了一点问题。
                MergedArea = Result.MergeArea;
            } while (MergedArea != null && MergedArea.Cells[1, 1].Address != FirstResult.Address);
            MessageChange(this,new MessageEventArgs($"搜寻完毕，共发现{t}处合并区域"));
        }
    }
    public class Url//这TM写得真够大的
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
            HttpClient checkClient = new HttpClient
            {
                Timeout = new TimeSpan(0, 0, 1)
            };
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
}
