using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;
namespace _1MathSetUp
{
    class Program
    {
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
        static void Main(string[] args)
        {
            //try
            //{
                CertAsync().Wait();
            //}
            //catch
            //{
            //    Console.WriteLine("错误：请联网并使用管理员权限运行");
            //    Console.ReadKey();
            //}
            TrustDirSetter.SetTrustDir();
            SetUpAsync().Wait();
            Console.ReadKey();
        }
        private static async Task CertAsync()
        {
            string localFullName = $"{DownloadPath}sn.cer";
            Uri uri = new Uri("http://Public.mujiannan.me/1Math/sn.cer");
            await Download(uri, localFullName);
            using (X509Certificate2 myCert = new X509Certificate2(localFullName))
            {
                using (X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine))
                {
                    store.Open(OpenFlags.ReadWrite);
                    //store.Remove(myCert);
                    store.Add(myCert);
                }
            }
        }
        private static async Task SetUpAsync()
        {
            string localFullName = $"{DownloadPath}setup.exe";
            Uri uri = new Uri("http://Public.mujiannan.me/1Math/setup.exe");
            Console.WriteLine($"下载安装包，将存放在{localFullName}");
            await Download(uri, localFullName);
            System.Diagnostics.Process.Start(localFullName).WaitForExit();
        }
        private static async Task Download(Uri uri, string localFullName)
        {
            HttpClient httpClient = new HttpClient();
            FileStream fileStream = File.Create(localFullName, 1024, FileOptions.Asynchronous);
            byte[] bytes = await httpClient.GetByteArrayAsync(uri);
            fileStream.Write(bytes,0,bytes.Length);//限制为100M，不知道0会不会是无限
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
