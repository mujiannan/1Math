using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

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

}
