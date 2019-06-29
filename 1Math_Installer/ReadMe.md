# 无商业证书场景下的VSTO分发方案

## 前言
本文旨在优化无商业证书场景下的VSTO分发方案，目标是实现用户端的简单安装、自动更新。

## 项目介绍
我花费几个月时间断断续续折腾出了一个Excel VSTO，曾在分发阶段遭遇了许多问题。
虽然麻烦不断，但通过不断努力（地搜索），最终得到了一套比较完善的方案，遂撰文分享一二。  
**主要思路**：在VSTO的ClickOnce之前套一层安装器，进行证书导入等操作，之后自动下载VSTO安装包、自动运行。  

1. 自制证书  
1. 给VSTO签名  
1. 发布VSTO   
1. 自制安装器  
   1. 安装器关键方法一：尝试下载自制证书，并导入用户机器  
   1. 安装器关键方法二：若证书导入失败，则尝试写入注册表  
   1. 安装器关键方法三：下载VSTO安装文件，运行  
1. 发布安装器

你可以看一下我最终实现的效果：[`安装测试`](
http://mujiannan.me/1Math/Installer/1Math_Installer.exe)[`开源地址`](
https://github.com/mujiannan/1Math)  
由于不是商业项目，我没有任何兼容性顾虑，项目中的.net Framework和Interop.Excel都选择了最新版本，因此可能会不兼容许多低版本系统或office。
但是，不用担心，事实上，即便顾及到兼容性，我的分发思路也应该是通用的。

## Excel VSTO部署简介
直接在Visual Studio中点击“发布”，会得到如下结构的文件（微软官方的图，使用了Outlook）。  
![publishfolderstructure](https://mujiannan.oss-cn-shanghai.aliyuncs.com/pictures/write/publishfolderstructure.png)  
这种发布方式被称为“ClickOnce”：[Deploy an Office solution by using ClickOnce - Visual Studio​docs.microsoft.com](
https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=vs-2017#Custom)  
**注：**
>*官网还介绍了以Windows Installer打包发布VSTO的方式，这种方式不支持自动升级。  
>一些第三方打包工具，如AdvancedInstaller，我不确定它们能否打包出可自动升级的VSTO。讲道理，它们“应该”能做到，毕竟它们是专业的嘛。*

不与如今的各种应用商店作比较，这种发布方式看起来是非常优秀的。
你可以指定用户从你的网站下载setup.exe，打开后安装器会自动从你设定的网络路径拉取所需文件，
安装完成之后它能够保持自动更新……这一切看起来都相当完美，但前提是你得有一个已经获取系统授信的**代码证书**。  
我也曾尝试从Comodo申请个人代码证书，但是，经过一阵沟通之后，个人代码证书似乎需要进行Face-to-Face验证，而我身在中国……

如果使用自己的测试证书之类的进行ClickOnce发布，并妄图将安装源设定在你的网站，甚至还妄想它能够保持自动更新……那么，你会得到这个：  
![截图来自网络](https://mujiannan.oss-cn-shanghai.aliyuncs.com/pictures/write/VSTOWithoutCert.png)  
当然，这张图是我在网上随便找的，我自己的分发问题已经彻底解决了，所以不会再出现这个界面。

## 网络上的方案

为了兼顾简易安装与自动更新，我翻阅了网络上的各种`Q&A`，
答案主要分为两派：“自制证书派”和“改注册表派”（以下分别简称“证书派”与“注册表派”）。  


**证书派：**自制代码证书，让用户导入你的证书，参见[通商软件MAX](
https://www.jianshu.com/p/db72e0c4545d?utm_campaign=maleskine&utm_content=note&utm_medium=seo_notes&utm_source=recommendation)的文章  
**注册表派：**注入注册表，修改用户机器的信任提示设置，参见[微软官方文档](
https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=vs-2017#Custom)

证书派的导入自制证书需要让用户多出一步操作，注册表派的注入注册表也需要让用户在安装VSTO之前运行一个事先写好的脚本。
## 我的分发方案

我的思路非常简单，如开头所述，
就仅仅是在VSTO的ClickOnce之前加一层安装器，进行证书导入等操作，之后自动下载VSTO安装包、自动运行。
1. 自制证书  
工具任选，自制一个代码证书，最终效果差不多是这样：  
![自制证书](https://mujiannan.oss-cn-shanghai.aliyuncs.com/pictures/write/DIYCodeCert.png)  
如果你不知道如何自制证书，那么可以从[这里](https://blog.csdn.net/tcjiaan/article/details/12394045)得到帮助。
1. 给VSTO签名  
![使用自制证书给VSTO签名](https://mujiannan.oss-cn-shanghai.aliyuncs.com/pictures/write/SignVSTO.png)
1. 发布VSTO  
虽然你使用了不受用户信任的自制证书，但你可以假装它是正规商业证书，直接把你的VSTO发布到站点上就是了。  
到此为止，你已经成功获得了一个**用户无法安装的VSTO**。它应该会有一个类似于“setup.exe”的文件，不要让用户直接下载它，它应该由你的自制安装器下载到用户电脑上。
1. 自制安装器  
它的名字可能会类似于“Installer.exe”，既可以是命令行脚本也可以是拥有华丽界面的WPF程序，这些都不重要，你可以随心所欲。  
重要的是，它应该是一个单文件软件，开袋即食的那种单文件软件，以方便用户下载、直接点击运行。  
它的作用主要有以下三点，我将以代码说明：
   1. 尝试下载自制证书，并导入用户机器
		
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
   1. 若证书导入失败，则尝试写入注册表
			
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
   1. 下载VSTO安装文件，运行  
这一步事实上是下载了VSTO-ClickOnce的setup.exe文件并自动运行
			
			private async Task SetUpAsync()
			{
				string localFullName = $"{DownloadPath}setup.exe";
				Uri uri = new Uri("http://Public.mujiannan.me/1Math/setup.exe");
				Reportor.Report($"下载{uri.AbsolutePath}至{localFullName}，请耐心等待...");
				await DownloadAsync(uri, localFullName);
				System.Diagnostics.Process.Start(localFullName);
			}
	下面是我制作的安装器的后台主方法：

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
1. 发布安装器  
你应该只把安装器发布给最终用户，用户在你的站点上或通过其它渠道下载安装器，这个安装器将如期运行：自动下载证书、自动导入证书、自动下载VSTO-ClickOnce的setup.exe文件、自动安装VSTO。
## 后记
作为一名非科班出身的编程初学者，一些像VSTO分发之类在专业开发者眼中可能根本不能算是问题的问题也会成为我的拦路虎。  
我相信专业的开发者们总能很自然地处理好这些事情，但我知道，还有一些跟我一样自学入门的开发者们在面对这些问题时也是头疼不已。
希望这篇文章可以帮助到一些人处理好无商业证书场景下的VSTO分发问题。  
学习路上，一帆风顺！