using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Net.Http;
using System.Windows.Media;

namespace _1Math
{
    internal class BackGroudTask
    {
        public BackGroudTask(ExcelConcurrent excelConcurrent)
        {

        }
    }
    internal abstract class ExcelConcurrent : IHasStatusReporter//定位：用于批量执行针对Excel.Range的任务
    {
        //source and results
        private Excel.Range _sourcesRange;
        private string[] _sources;
        private int _m;
        private int _n;
        private volatile dynamic[] _results;

        //concurrent controller and progress reportor
        int _maxConcurrent;
        public event ChangeMessage MessageChange;
        public event ChangeProgress ProgressChange;
        private volatile int _totalCount;
        private volatile int _completedCount=0;
        private void CompleteOneTask()
        {
            _completedCount++;
            ProgressChange(this, new ProgressEventArgs(_completedCount / (double)_totalCount));
            MessageChange(this, new MessageEventArgs($"已处理：{_completedCount}/{_totalCount}"));
        }

        private void Initialize(Excel.Range range)
        {
            _sourcesRange = range;
            _m = _sourcesRange.Rows.Count;
            _n = _sourcesRange.Columns.Count;
            _totalCount = _m * _n;
            _maxConcurrent = Math.Min(System.Environment.ProcessorCount, _totalCount);
        }
        public ExcelConcurrent()
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (range == null)
            {
                throw new Exception("SelectionIs'ntExcelRange");
            }
            else if (range.Areas.Count > 1)
            {
                throw new Exception("SelectionIs'tContinuousExcelRange");//不连续区域……竖的还好，横的就搞笑了，直接不允许用户这么干才是最好的
            }
            else
            {
                Initialize(range);
            }
        }
        public async Task StartAsync(CancellationToken cancellationToken)
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            MessageChange(this, new MessageEventArgs("初始化任务资源..."));
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

            //Wait for workers to finish
            await Task.WhenAll(tasks);

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
            if (!cancellationToken.IsCancellationRequested)
            {
                _sourcesRange.Offset[0, _n].Value=results;
            }
            stopwatch.Stop();
            MessageChange(this, new MessageEventArgs($"耗时{stopwatch.Elapsed.TotalSeconds}秒，{_completedCount}/{_totalCount}"));
        }

        //任务分配
        private const int NoNext = -1;
        private volatile int _builtCount=0;
        private int GetNext()
        {
            int next;
            if (_builtCount<_totalCount)
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
            while (next!=NoNext)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    return;
                }
                string source = _sources[next];
                _results[next] = await WorkAsync(source, cancellationToken);
                CompleteOneTask();
                next = GetNext();//2019年6月7日脑残了，忘了加这句……
            }
        }
        protected abstract Task<dynamic> WorkAsync(string source,CancellationToken cancellationToken);
    }
    internal sealed class AccessibilityChecker:ExcelConcurrent
    {
        protected override async Task<dynamic> WorkAsync(string source, CancellationToken cancellationToken)
        {
            string url =(string)source;
            using (HttpClient checkClient = new HttpClient())
            {
                try
                {
                    var response = await checkClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
                    return response.IsSuccessStatusCode.ToString();
                }
                catch (Exception Ex)
                {
                    return (Ex.Message);
                }
            }
        }
    }
    internal sealed class MediaDurationChecker : ExcelConcurrent
    {
        protected override async Task<dynamic> WorkAsync(string source, CancellationToken cancellationToken)
        {
            double duration;
            duration = await Task.Run(delegate
            {
                MediaPlayer mediaPlayer = new MediaPlayer();
                mediaPlayer.Open(new Uri(source));
                double result=0;
                DateTime start = DateTime.Now;
                TimeSpan timeSpan;
                do
                {
                    Thread.Sleep(50);
                    if (mediaPlayer.NaturalDuration.HasTimeSpan)
                    {
                        result = mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds;
                        mediaPlayer.Close();
                    }
                    timeSpan = DateTime.Now - start;
                } while (result == 0 && timeSpan.TotalSeconds < 10);
                return (result);
            });
            return duration;
        }
    }
}
