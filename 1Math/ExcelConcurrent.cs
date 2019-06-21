using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Net.Http;
using System.Windows.Media;
using QRCoder;
namespace _1Math
{
    internal abstract class ExcelConcurrent : IReportor//定位：用于批量执行针对Excel.Range的任务
    {
        //source and results
        private Excel.Range _sourcesRange;
        private string[] _sources;
        private int _m;
        private int _n;
        private dynamic[] _results;
        public int[] ResultOffSet { get; set; } = new int[2] { 0, 1 };

        //concurrent controller and progress reportor
        int _maxConcurrent = 2;
        private int MaxConcurrent
        {
            get => _maxConcurrent;
            set
            {
                if (value > 0)
                {
                    _maxConcurrent = Math.Min(value, _totalCount);
                    //这里的逻辑也要处理好，尽量不让开启的并发任务数大于总任务数，避免资源浪费
                    //当然，就算大于了，后面的GetNext方法也会让多余的Task安安静静地结束，不会有任何bug
                }
                else
                {
                    _maxConcurrent = Math.Min(System.Environment.ProcessorCount, _totalCount);
                    //输入零就自动确定线程数
                }
            }
        }
        public Reportor Reportor { get; private set; }
        private volatile int _totalCount;
        private volatile int _completedCount = 0;
        private void CompleteOneTask()
        {
            _completedCount++;
            Reportor.Report(_completedCount / (double)_totalCount);
            Reportor.Report($"已处理：{_completedCount}/{_totalCount}");
        }
        public ExcelConcurrent(Excel.Range range = null, int maxConcurrent = 0)
        {
            if (range == null)//子类省略了range参数，则默认为当前selection
            {
                range = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (range == null)//如果使用selection，则需要检测一下它是否为range，结合上面的as判断
                {
                    throw new Exception("InputIs'ntExcelRange");
                }
            }
            if (range.Areas.Count > 1)
            {
                throw new Exception("DisontinuousExcelRange");//不连续区域……竖的还好，横的就搞笑了，直接不允许用户这么干才是最好的
            }
            _sourcesRange = range;
            _m = _sourcesRange.Rows.Count;
            _n = _sourcesRange.Columns.Count;
            _totalCount = _m * _n;
            MaxConcurrent = maxConcurrent;
            Reportor = new Reportor(this);
        }
        public async Task StartAsync(CancellationToken cancellationToken)
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            Reportor.Report("初始化任务资源...");
            _sources = new string[_totalCount];
            int t = 0;//
            try
            {
                for (int i = 0; i < _m; i++)
                {
                    for (int j = 0; j < _n; j++)
                    {
                        _sources[t] = (string)_sourcesRange[i + 1, j + 1].Value;
                        t++;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
#if DEBUG
            string t1 = stopwatch.Elapsed.TotalSeconds.ToString();
#endif
            //Build tasks
            _results = new dynamic[_totalCount];

            Task[] tasks = new Task[_maxConcurrent];
            for (int i = 0; i < _maxConcurrent; i++)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    return;
                }
                tasks[i] = WorkContinuouslyAsync(cancellationToken);
            }
#if DEBUG
            string t2 = stopwatch.Elapsed.TotalSeconds.ToString();
#endif
            //Wait for workers to finish
            await Task.WhenAll(tasks);
            string t3 = stopwatch.Elapsed.TotalSeconds.ToString();
            //Get results
            dynamic[,] results = new dynamic[_m, _n];
            t = 0;
            for (int i = 0; i < _m; i++)
            {
                for (int j = 0; j < _n; j++)
                {
                    results[i, j] = _results[t];
                    t++;
                }
            }
#if DEBUG
            string t4 = stopwatch.Elapsed.TotalSeconds.ToString();
#endif
            await WriteBackAsync(results, cancellationToken);
            stopwatch.Stop();
            Reportor.Report($"耗时{stopwatch.Elapsed.TotalSeconds}秒，{_completedCount}/{_totalCount}");
#if DEBUG
            if (cancellationToken.IsCancellationRequested)
            {
                System.Windows.Forms.MessageBox.Show($"已取消，耗时{t4}");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show($"从Excel读完数据{t1} 构建完任务{t2} 执行完任务{t3} 回写完Excel{t4}");
            }
#endif
        }

        //2019年6月12日 为了兼容多功能，拆分出回写方法并使之可覆盖
        protected virtual async Task WriteBackAsync(dynamic[,] results, CancellationToken cancellationToken)//等着覆盖
        {
            await Task.Run(() =>
            {
                if (!cancellationToken.IsCancellationRequested)
                {
                    _sourcesRange.Offset[_m * ResultOffSet[0], _n * ResultOffSet[1]].Value = results;
                }
            });
        }

        //任务分配
        private const int NoNext = -1;
        private volatile int _builtCount = 0;
        private int GetNext()
        {
            int next;
            if (_builtCount < _totalCount)
            {
                next = _builtCount;
                _builtCount++;
            }
            else
            {
                next = NoNext;
            }
            return next;
        }

        //持续工作的线程
        private async Task WorkContinuouslyAsync(CancellationToken cancellationToken)
        {
            int next = GetNext();
            while (next != NoNext)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    return;
                }
                string source = _sources[next];
                _results[next] = await WorkAsync(source, next, cancellationToken);
                CompleteOneTask();
                next = GetNext();//2019年6月7日脑残了，忘了加这句……
            }
        }
        protected abstract Task<dynamic> WorkAsync(string source, int sourceID, CancellationToken cancellationToken);
    }
    internal sealed class AccessibilityChecker : ExcelConcurrent
    {
        public AccessibilityChecker() : base(null, Environment.ProcessorCount * 4) { }
        protected override async Task<dynamic> WorkAsync(string source, int sourceID, CancellationToken cancellationToken)
        {
            string url = (string)source;
            string accessibility;
            using (HttpClient checkClient = new HttpClient())
            {
                HttpResponseMessage response;
                try
                {
                    response = await checkClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
                    accessibility = response.IsSuccessStatusCode.ToString();
                    response.Dispose();//Must dispose it, otherwise the internet will run out of bandwidth
                }
                catch (Exception Ex)
                {
                    accessibility = Ex.Message;
                }
            }
            return accessibility;
        }
    }
    internal sealed class MediaInfoChecker : ExcelConcurrent
    {
        public bool CheckDuration { get; set; } = false;
        public bool CheckHasVideo { get; set; } = false;
        public bool CheckHasAudio { get; set; } = false;
        public bool CheckResolution { get; set; } = false;
        private byte CheckItemsCount
        {
            get
            {
                byte count = 0;
                if (CheckDuration) count++;
                if (CheckHasVideo) count++;
                if (CheckHasAudio) count++;
                if (CheckResolution) count++;
                return count;
            }
        }


        protected override async Task<dynamic> WorkAsync(string source, int sourceID, CancellationToken cancellationToken)
        {
            if (CheckItemsCount == 0)
            {
                throw new Exception("CheckItemsCount==0");
            }
            double duration = 0;
            bool hasVideo = false;
            bool hasAudio = false;
            string resolution = string.Empty;
            await Task.Run(() =>
            {
                MediaPlayer mediaPlayer = new MediaPlayer
                {
                    ScrubbingEnabled = true
                };
                mediaPlayer.Open(new Uri(source));//胡乱输入的话，Debug阶段会有异常，Release版本没问题，最外层会处理好
                DateTime start = DateTime.Now;
                TimeSpan timeSpan;
                try
                {
                    do
                    {
                        Thread.Sleep(50);
                        if (mediaPlayer.NaturalDuration.HasTimeSpan)
                        {
                            mediaPlayer.Pause();
                            duration = mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                            hasVideo = mediaPlayer.HasVideo;
                            hasAudio = mediaPlayer.HasAudio;
                            resolution = mediaPlayer.NaturalVideoWidth.ToString() + "*" + mediaPlayer.NaturalVideoHeight.ToString();
                        }
                        timeSpan = DateTime.Now - start;
                    } while (duration == 0 && timeSpan.TotalSeconds < 100);
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    mediaPlayer.Stop();
                    mediaPlayer.Close();
                    mediaPlayer = null;
                }
            });
            StringBuilder result = new StringBuilder(16);
            if (CheckDuration)
            {
                result.Append(duration);
            }
            if (CheckHasVideo)
            {
                if (result.Length > 0) result.Append(",");
                if (hasVideo)
                {
                    result.Append("HasVideo");
                }
                else
                {
                    result.Append("NoVideo");
                }
            }
            if (CheckHasAudio)
            {
                if (result.Length > 0) result.Append(",");

                if (hasAudio)
                {
                    result.Append("HasAudio");
                }
                else
                {
                    result.Append("NoAudio");
                }
            }
            if (CheckResolution)
            {
                if (result.Length > 0) result.Append(",");
                result.Append(resolution);
            }
            return result.ToString();
        }
    }
    internal sealed class QRGenerator : ExcelConcurrent
    {
        internal QRGenerator(string path)
        {
            _path = path;
        }
        private string _path;
        protected override async Task<dynamic> WorkAsync(string source, int sourceID, CancellationToken cancellationToken)
        {
            System.Drawing.Bitmap bitmap;
            Task<System.Drawing.Bitmap> task = Task.Run(() =>
            {
                using (QRCodeGenerator qRCodeGenerator = new QRCodeGenerator())
                {
                    using (QRCodeData qRCodeData = qRCodeGenerator.CreateQrCode(source, QRCodeGenerator.ECCLevel.H))
                    {
                        using (QRCode qRCode = new QRCode(qRCodeData))
                        {
                            return qRCode.GetGraphic(20);
                        }
                    }
                }
            });
            bitmap = await task;
            string fullName = _path + "\\" + sourceID + ".jpeg";
            bitmap.Save(fullName, System.Drawing.Imaging.ImageFormat.Jpeg);
            bitmap.Dispose();
            return fullName;
        }
    }
}
