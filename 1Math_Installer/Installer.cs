using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;
using System.Net.Http.Formatting;
using System.Net.Http.Handlers;
using Correspondence;

namespace _1Math_Installer
{
    

    class Installer: IReportor
    {
        public Reportor Reportor { get; }//它只在构造函数中初始化，之后就是只读的（当然，可以令其向外报告事件）
        public Installer()
        {
            this.Reportor = new Reportor(this);
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
                string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData) + @"\1Math\Downloads\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                return path;
            }
        }
        
        public async Task StartInstallerAsync()
        {
            try
            {
                await CertAsync();
                TrustDirSetter.SetTrustDir();
            }
            catch
            {
                throw new Exception("错误：请联网并使用管理员权限运行");
            }
            await SetUpAsync();
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

        private static class TrustDirSetter
        {
            private static RegistryKey Key
            {
                get
                {
                    if (System.Environment.Is64BitOperatingSystem)
                    {
                        return (Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                            @"SOFTWARE\Wow6432Node\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel"));
                    }
                    else
                    {
                        return (Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                            @"SOFTWARE\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel"));
                    }
                }
            }
            internal static void SetTrustDir()
            {
                Key.SetValue("MyComputer", "Disabled");
                Key.SetValue("LocalIntranet", "Disabled");
                Key.SetValue("Internet", "Disabled");
                Key.SetValue("TrustedSites", "Disabled");
                Key.SetValue("UntrustedSites", "AuthenticodeRequired");
                Key.Close();
            }
        }

    }

}
