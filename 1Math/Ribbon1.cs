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
            CommonExcel commonExcel = new CommonExcel();
            NetTask netTask = new NetTask();
            netTask.CheckUrlsAccessibility(commonExcel.SelectedRange);
        }

        private void buttonVideoLength_Click(object sender, RibbonControlEventArgs e)
        {
            CommonExcel commonExcel = new CommonExcel();
            NetTask netTask = new NetTask();
            netTask.CheckVideosLength(commonExcel.SelectedRange);
        }

        private void ButtonTest_Click(object sender, RibbonControlEventArgs e)
        {
            Test test = new Test();
            test.TestIt();
        }
    }
}
