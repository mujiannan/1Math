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
        private void BuildNetTask()
        {
            tasks = new Tasks();
            tasks.MessageChange += NetTask_MessageChange;
            tasks.SheduleChange += NetTask_SheduleChange;
        }
        System.Threading.Thread NetTaskThread;
        private Tasks tasks;
        private delegate void DBuildNetTaskThread();
        private DBuildNetTaskThread dBuildNetTaskThread;
        private void StartNetTaskThread()
        {
            if (NetTaskThread == null)
            {
                NetTaskThread = new System.Threading.Thread(dBuildNetTaskThread.Invoke);
                NetTaskThread.Start();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请等待其它NetTask完成");
            }
        }
        private void ButtonUrlCheck_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStatusForm();
            BuildNetTask();
            dBuildNetTaskThread = new DBuildNetTaskThread(tasks.CheckUrlsAccessibility);
            StartNetTaskThread();
        }
        private void ButtonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStatusForm();
            BuildNetTask();
            dBuildNetTaskThread = new DBuildNetTaskThread(tasks.CheckVideosLength);
            StartNetTaskThread();
        }
        private void StatusForm_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            if (NetTaskThread!=null)
            {
                if (NetTaskThread.IsAlive)
                {
                    NetTaskThread.Abort();
                }
                NetTaskThread = null;
            }
        }
        private void ShowStatusForm()
        {
            statusForm = new StatusForm();
            statusForm.Show();
            statusForm.FormClosing += StatusForm_FormClosing;
        }
        private void NetTask_SheduleChange(double NewStatus)
        {
            statusForm.progressBar1.Invoke(new Action(() => { statusForm.progressBar1.Value = (int)(100 * NewStatus); }));
        }
        private void NetTask_MessageChange(string NewStatus)
        {
            statusForm.MessageLabel.Invoke(new Action(() => { statusForm.MessageLabel.Text = NewStatus; })); ;
        }

        private void ButtonTest_Click(object sender, RibbonControlEventArgs e)
        {
            Test test = new Test();
            test.TestIt();
        }

        private void ButtonAntiMerge_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStatusForm();
            BuildNetTask();
            dBuildNetTaskThread = new DBuildNetTaskThread(tasks.AntiMerge);
            StartNetTaskThread();
        }
    }
}
