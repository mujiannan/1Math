using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;

namespace _1Math
{
    public class Url//这TM写得真够大的
    {
        private string str;
        public void SetReferTo(string value)
        {
            if (value != str)
            {
                checkStatus = CheckStatus.Null;//这样可以不需要反复实例化就可以验证有效性，不知道会不会节省资源
                str = value;
            }
        }
        public new string ToString()//假的
        {
            return str;
        }
        public Url(string url)
        {
            str = url;
        }
        public Url()
        {
        }
        private enum CheckStatus
        {
            Null, Checking, Checked
        }
        private Task checkTask;
        private CheckStatus checkStatus;
        public void CheckAccessibility()
        {
            checkTask = Check();
        }
        private async Task Check()
        {
            checkStatus = CheckStatus.Checking;
            HttpClient checkClient = new HttpClient
            {
                Timeout = new TimeSpan(0, 0, 1)
            };
            try
            {
                HttpResponseMessage httpResponseMessage = await checkClient.GetAsync(str, HttpCompletionOption.ResponseHeadersRead);
                accessibility = httpResponseMessage.IsSuccessStatusCode;
            }
            catch (Exception)
            {
                accessibility = false;
            }
            checkClient.Dispose();
            checkStatus = CheckStatus.Checked;
        }
        private bool accessibility;
        public bool Accessibility
        {
            get
            {
                switch (checkStatus)
                {
                    case CheckStatus.Null:
                        CheckAccessibility();
                        checkTask.Wait();
                        break;
                    case CheckStatus.Checking:
                        checkTask.Wait();
                        break;
                    case CheckStatus.Checked:
                        break;
                    default:
                        break;
                }
                return accessibility;
            }
        }

    }
}
