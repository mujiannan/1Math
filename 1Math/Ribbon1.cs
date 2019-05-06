using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
namespace _1Math
{
    public partial class Ribbon1
    {
        System.Threading.Tasks.Task backGroundTask;
        System.Threading.Tasks.Task BackGroundTask
        {
            set
            {
                if (backGroundTask==null)
                {
                    backGroundTask = value;
                }
                else
                {
                    throw (new Exception("AnotherBackGroundTaskIsRunning"));
                }
            }
        }
        
        private StatusForm statusForm;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void ButtonUrlCheck_Click(object sender, RibbonControlEventArgs e)
        {
            if (statusForm != null)
            {
                System.Windows.Forms.MessageBox.Show("AnotherBackGroundTaskIsRunning", "PleaseWaitForThePreviousTask");
                return;
            }
            ShowStatusForm();
            CTS = new System.Threading.CancellationTokenSource();
            Accessibility accessibility = new Accessibility();
            accessibility.StatusChange += _StatusChange;
            BackGroundTask = new System.Threading.Tasks.Task(() =>
              {
                  accessibility.Start(CTS.Token);
              }, CTS.Token);
            backGroundTask.Start();
        }
        System.Threading.CancellationTokenSource CTS;
        private void ButtonAntiMerge_Click(object sender, RibbonControlEventArgs e)
        {
            if (statusForm!=null)
            {
                System.Windows.Forms.MessageBox.Show("AnotherBackGroundTaskIsRunning","PleaseWaitForThePreviousTask");
                return;
            }
            ShowStatusForm();
            CTS=new System.Threading.CancellationTokenSource();
            MergeAreas mergeAreas = new MergeAreas();
            mergeAreas.StatusChange += _StatusChange;
            BackGroundTask = new System.Threading.Tasks.Task(() =>
            {
                mergeAreas.SafelyUnMergeAndFill(CTS.Token);
            }, CTS.Token);
            backGroundTask.Start();
        }
        private void _StatusChange(object sender, Status NewStatus)
        {
            if (statusForm!=null)
            {
                statusForm.BeginInvoke(new Action(() => { statusForm.progressBar1.Value = (int)(100 * NewStatus.Progress); }));
                statusForm.BeginInvoke(new Action(() => { statusForm.MessageLabel.Text = NewStatus.Message; }));
            }
        }

        private void ButtonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            if (statusForm != null)
            {
                System.Windows.Forms.MessageBox.Show("AnotherBackGroundTaskIsRunning", "PleaseWaitForThePreviousTask");
                return;
            }
            ShowStatusForm();
            CTS = new System.Threading.CancellationTokenSource();
            VideoLength videoLength = new VideoLength();
            videoLength.StatusChange += _StatusChange;
            BackGroundTask = new System.Threading.Tasks.Task(() =>
              {
                  videoLength.Start(CTS.Token);
              }, CTS.Token);
            backGroundTask.Start();
        }


        private void StatusForm_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            if (!CTS.IsCancellationRequested)
            {
                CTS.Cancel();
            }
            CTS.Dispose();
            CTS = null;
            backGroundTask.Dispose();
            backGroundTask = null;
            CE.EndTask();
            statusForm.Dispose();
            statusForm = null;
        }
        private void ShowStatusForm()
        {
            statusForm = new StatusForm();
            statusForm.Show();
            statusForm.FormClosing += StatusForm_FormClosing;
        }
    }
}
