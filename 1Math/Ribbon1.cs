using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
namespace _1Math
{
    public partial class Ribbon1
    {
        private StatusForm statusForm;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void BuildTask()
        {
            tasks = new Tasks();
            tasks.MessageChange += Task_MessageChange;
            tasks.ProgressChange += Task_ProgressChange;
        }
        System.Threading.Thread TaskThread;
        private Tasks tasks;
        private delegate void DBuildTaskThread();
        private DBuildTaskThread dBuildTaskThread;
        private void StartTaskThread()
        {
            if (TaskThread == null)
            {
                TaskThread = new System.Threading.Thread(dBuildTaskThread.Invoke);
                TaskThread.Start();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请等待其它NetTask完成");
            }
        }
        private void ButtonUrlCheck_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStatusForm();
            BuildTask();
            dBuildTaskThread = new DBuildTaskThread(tasks.CheckUrlsAccessibility);
            StartTaskThread();
        }
        private void ButtonAntiMerge_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStatusForm();
            BuildTask();
            dBuildTaskThread = new DBuildTaskThread(tasks.AntiMerge);
            StartTaskThread();
        }
        private void ButtonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStatusForm();
            BuildTask();
            dBuildTaskThread = new DBuildTaskThread(tasks.CheckVideosLength);
            StartTaskThread();
        }
        private void StatusForm_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            if (TaskThread!=null)
            {
                if (TaskThread.IsAlive)
                {
                    TaskThread.Abort();
                }
                TaskThread = null;
            }
        }
        private void ShowStatusForm()
        {
            statusForm = new StatusForm();
            statusForm.Show();
            statusForm.FormClosing += StatusForm_FormClosing;
        }
        private void Task_ProgressChange(double NewStatus)
        {
            statusForm.progressBar1.BeginInvoke(new Action(() => { statusForm.progressBar1.Value = (int)(100 * NewStatus); }));
        }
        private void Task_MessageChange(string NewStatus)
        {
            statusForm.MessageLabel.BeginInvoke(new Action(() => { statusForm.MessageLabel.Text = NewStatus; })); ;
        }
    }
}
