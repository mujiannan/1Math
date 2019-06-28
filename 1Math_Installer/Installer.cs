using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Http.Handlers;
using Correspondence;
using Microsoft.Win32;

namespace _1Math_Installer
{
    class Installer: IReportor
    {
        public Reportor Reportor { get; }
        public Installer()
        {
            this.Reportor = new Reportor(this);
        }
        public async Task StartInstallerAsync()
        {
            //进行两步操作（导入证书、注入注册表）
            //优先选择导入证书，如果证书导入失败，则转而尝试注册表
            //当两步操作都失败时，向外传递异常信息

            string errInfo = string.Empty;//用于记录错误信息
            bool trustSuccess = false;//标记证书是否导入成功
            try
            {
                await CertAsync();//导入证书
                trustSuccess = true;
            }
            catch (Exception Ex)
            {
                errInfo += Ex.Message;//记录证书导入错误消息
                Reportor.Report("错误：证书导入失败，请联网并使用管理员权限运行");
            }
            if (!trustSuccess)
            {
                try//事实上，如果上一步证书没导入成功，再来尝试写注册表也基本上是无济于事的
                {
                    Reg.SetTrustPromptBehavior();//写入注册表
                    trustSuccess = true;
                }
                catch (Exception Ex)
                {
                    errInfo += Ex.Message;
                    this.Reportor.Report("错误：注册表写入失败，请尝试使用管理员权限运行");
                }
            }
            if (trustSuccess)//当导入证书与注册表写入有任何一项成功时，都可以正常执行下载、安装的操作
            {
                await SetUpAsync();
            }
            else
            {
                await Task.Delay(2000);//显示最终的错误信息前，留点时间告诉用户要用管理员权限运行
                throw new Exception(errInfo);//别在本地存日志之类的，直接显示出来吧……
            }
        }
        private static string Local1MathPath
        {
            get
            {
                string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData) + @"\1Math\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                return path;
            }
        }
        private static string DownloadPath
        {
            get
            {
                string path = Local1MathPath + @"\Downloads\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                return path;
            }
        }
        
        private async Task CertAsync()
        {
            string localFullName = $"{DownloadPath}sn.cer";
            Uri uri = new Uri("http://Public.mujiannan.me/1Math/sn.cer");
            await DownloadAsync(uri, localFullName);
            using (X509Certificate2 myCert = new X509Certificate2(localFullName))
            {
                using (X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine))
                {
                    store.Open(OpenFlags.ReadWrite);
                    store.Remove(myCert);
                    store.Add(myCert);
                }
            }
        }
        private async Task SetUpAsync()
        {
            string localFullName = $"{DownloadPath}setup.exe";
            Uri uri = new Uri("http://Public.mujiannan.me/1Math/setup.exe");
            Reportor.Report($"下载{uri.AbsolutePath}至{localFullName}，请耐心等待...");
            await DownloadAsync(uri, localFullName);
            System.Diagnostics.Process.Start(localFullName);
        }
        private async Task DownloadAsync(Uri uri, string localFullName)
        {
            
            HttpClientHandler httpClientHandler = new HttpClientHandler();
            
            ProgressMessageHandler progressMessageHandler = new ProgressMessageHandler(httpClientHandler);
            HttpClient httpClient = new HttpClient(progressMessageHandler);
            progressMessageHandler.HttpReceiveProgress += (object sender, HttpProgressEventArgs e) => Reportor.Report(e.ProgressPercentage);//捕获下载进度，向外汇报
            FileStream fileStream = File.Create(localFullName, 1024, FileOptions.Asynchronous);
            byte[] bytes = await httpClient.GetByteArrayAsync(uri);
            fileStream.Write(bytes, 0, bytes.Length);
            fileStream.Flush();
            fileStream.Close();
        }
        //导入证书与设置注册是两个可用的选项
        //当可以导入证书时，最好不要去动注册表
        private static class Reg
        {

            internal static void SetTrustPromptBehavior()
            {
                RegistryKey key;
                if (System.Environment.Is64BitOperatingSystem)
                {
                    key=Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                        @"SOFTWARE\Wow6432Node\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel");
                }
                else
                {
                    key=Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                        @"SOFTWARE\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel");
                }
                key.SetValue("MyComputer", "Enabled");
                key.SetValue("LocalIntranet", "Enabled");
                key.SetValue("Internet", "Enabled");
                key.SetValue("TrustedSites", "Enabled");
                key.SetValue("UntrustedSites", "Enabled");
                key.Close();
            }
        }

    }

}
