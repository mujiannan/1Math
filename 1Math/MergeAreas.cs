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
    class MergeAreas//实例化后，直接运行SafelyUnMergedAndFill方法即可。拆分的目标默认为当前Selection。目标区域为Single时，则自动将目标区域更改为整个活动工作表
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
            _application = Globals.ThisAddIn.Application;
            _target = _application.Selection;
        }
        public MergeAreas(Excel.Range In)//此重载提供了将这个类用于快速获取合并单元格区域的可能性
        {
            _application = In.Application;
            _target = In;
        }
        public ArrayList ToArrayList()//类本身当然不能作为数组，但我可以为其添加ToArrayList方法，伪装一下
        {
            if (_mergedAreas == null)
            {
                GetMergedAreas();
            }
            ProgressChange(this, new ProgressEventArgs(1));
            return _mergedAreas;
        }
        public async void SafelyUnMergeAndFill(CancellationToken cancellationToken = new CancellationToken())
        {
            _cancellationToken = cancellationToken;
            if (_mergedAreas == null)
            {
                await Task.Run(new Action(GetMergedAreas));
                if (_mergedAreas == null)
                {
                    MessageChange(this, new MessageEventArgs("找不到合并的单元格"));
                    ProgressChange(this, new ProgressEventArgs(1));
                    return;
                }
            }
            ProgressChange(this, new ProgressEventArgs(0.5));
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
                ProgressChange(this, new ProgressEventArgs(0.5 + t / Sum / 2));
                range.Value = range.Cells[1, 1];//为什么这也能迭代……啥原因呢，不是说foreach不能这么来么
            }
            ProgressChange(this, new ProgressEventArgs(1));
        }
        private void GetMergedAreas()
        {

            _mergedAreas = new ArrayList();
            if (_target.Count == 1)
            {
                MessageChange(this, new MessageEventArgs("只选择了一个单元格，自动将搜寻区域拓展至其所在的整张工作表"));
                _target = _target.Worksheet.UsedRange;//这样的设定会使我们开发出更便于使用的VSTO
            }
            _application.FindFormat.MergeCells = true;
            Excel.Range Result = _target.Find(What: "", After: _target.Cells[1, 1], SearchFormat: true);
            Excel.Range FirstResult = Result;
            Excel.Range MergedArea = Result;
            if (FirstResult == null)
            {
                MessageChange(this, new MessageEventArgs("没有发现合并单元格！"));
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
            MessageChange(this, new MessageEventArgs($"搜寻完毕，共发现{t}处合并区域"));
        }
    }
}
