using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace _1Math
{
    class MergeAreas:IReportor//实例化后，直接运行SafelyUnMergedAndFill方法即可。拆分的目标默认为当前Selection。目标区域为Single时，则自动将目标区域更改为整个活动工作表
    {
        CancellationToken _cancellationToken;
        private Excel.Range _target;
        private ArrayList _mergedAreas;
        private Excel.Application _application;
        public Reportor Reportor { get; private set; }
        private void Initialize()
        {
            if (_target==null)
            {
                throw new Exception("PleaseGiveMeAnExcelRange");
            }
            if (_target.Areas.Count > 1)
            {
                throw new Exception("DiscontinuousExcelRange");
            }
            _application = _target.Application;
            Reportor = new Reportor(this);
        }
        public MergeAreas()
        {
            _target = Globals.ThisAddIn.Application.Selection as Excel.Range;//默认的当然是Selection
            Initialize();
        }
        public MergeAreas(Excel.Range In)
            //此重载提供了将这个类用于快速获取合并单元格区域的可能性
            //让来源可以是输入的其它区域而不限于Selection
        {
            _target = In;
            Initialize();
        }
        public async Task<ArrayList> ToArrayListAsync()
            //类本身当然不能作为ArrayList，但我可以为其添加ToArrayList方法，伪装一下，输出所有合并区域
        {
            if (_mergedAreas == null)
            {
                await Task.Run(GetMergedAreas);
            }
            Reportor.Report(1);
            return _mergedAreas;
        }
        public async Task SafelyUnMergeAndFill(CancellationToken cancellationToken = new CancellationToken())
        {
            _cancellationToken = cancellationToken;
            if (_mergedAreas == null)
            {
                Task findMergeAreas=new Task(GetMergedAreas, TaskCreationOptions.LongRunning);
                findMergeAreas.Start();
                await findMergeAreas;
                if (_mergedAreas == null)
                {
                    Reportor.Report("找不到合并的单元格");
                    Reportor.Report(1);
                    return;
                }
            }
            Reportor.Report(0.5);
            if (_cancellationToken.IsCancellationRequested)
            {
                return;
            }
            Reportor.Report("正在取消合并...");
            Reportor.Report(0.55);
            await Task.Run(_target.UnMerge);
            int t = 0;
            double Sum = _mergedAreas.Count;//为了进度报告不写强制转换，直接double吧
            Task fill=new Task(() =>
            {
                foreach (Excel.Range range in _mergedAreas)
                {
                    if (_cancellationToken.IsCancellationRequested)
                    {
                        return;
                    }
                    t++;
                    Reportor.Report($"取消合并中，第{t}个……");
                    Reportor.Report(0.55 + 0.45*t / Sum );
                    range.Value = range.Cells[1, 1];//为什么这也能迭代……啥原因呢，不是说foreach不能这么来么
                }
            },TaskCreationOptions.LongRunning);
            fill.Start();
            await fill;
            Reportor.Report($"大功告成！一共解决了{Sum}块合并区域！");
            Reportor.Report(1);
        }
        private void GetMergedAreas()
        {
            _mergedAreas = new ArrayList();
            if (_target.Count == 1)
            {
                Reportor.Report("只选择了一个单元格，自动将搜寻区域拓展至其所在的整张工作表");
                _target = _target.Worksheet.UsedRange;//这样的设定会使我们开发出更便于使用的VSTO
            }
            _application.FindFormat.MergeCells = true;
            Excel.Range Result = _target.Find(What: "", After: _target.Cells[1, 1], SearchFormat: true);
            Excel.Range FirstResult = Result;
            if (FirstResult == null)
            {
                Reportor.Report("没有发现合并单元格！");
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
            Excel.Range MergedArea = FirstResult.MergeArea;
            int t = 0;
            do
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    return;
                }
                t++;
                Reportor.Report($"搜寻中，已找到{t}处合并区域，总进度未知...");
                _mergedAreas.Add(MergedArea);
                Result = _target.Find(What: "", After: Result, SearchFormat: true); ;//这里的接龙很巧妙，但也很坑。我还尝试着用FindNext，但是出现了一点问题。
                MergedArea = Result.MergeArea;
            } while (MergedArea != null && MergedArea.Cells[1, 1].Address != FirstResult.Address);
            Reportor.Report($"搜寻完毕，共发现{t}处合并区域");
        }
    }
}
