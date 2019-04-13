using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
namespace _1Math
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ButtonUrlCheck_Click(object sender, RibbonControlEventArgs e)
        {
            NetTask netTask = new NetTask();
            System.Threading.Thread thread = new System.Threading.Thread(netTask.CheckUrlsAccessibility);
            thread.Start();
        }

        private void ButtonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            NetTask netTask = new NetTask();
            System.Threading.Thread thread = new System.Threading.Thread(netTask.CheckVideosLength);
            thread.Start();
        }

        private void ButtonTest_Click(object sender, RibbonControlEventArgs e)
        {
            Test test = new Test();
            test.TestIt();
        }
    }
}
