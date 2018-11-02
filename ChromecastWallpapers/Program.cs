using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Forms;
using Windows.Native;
using ChromecastWallpapers.Properties;
using OpenQA.Selenium.Chrome;
using Timer = System.Timers.Timer;

namespace ChromecastWallpapers
{
	public class Program
	{
		private static string _chromeDriverExe = "chromedriver.exe";
		private static ChromeDriver _driver;
		private static string _currentImageUrl;
		private static readonly Timer Timer = new Timer();

		[STAThread]
		static void Main(string[] args)
		{
			ExtractChromeDriver();
			void ChangeWallpaperInStaThread()
			{
				Thread thread = new Thread(ChangeWallPaper);
				thread.SetApartmentState(ApartmentState.STA);
				thread.Start();
				thread.Join(); 
			}
			shlobj.EnableActiveDesktop();
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			CreateDriver();
			Timer.Interval = 5000;
			Timer.Elapsed += (sender, eventArgs) =>
			{
				if (!IsDriverValid())
				{
					CreateDriver();
				}
				ChangeWallpaperInStaThread();
			};
			Timer.Start();
			ChangeWallPaper();
			Application.Run(new MyCustomApplicationContext());
		}

		private static void CreateDriver()
		{
			var options = new ChromeOptions();
			options.AddArguments("headless");
			ChromeDriverService service = ChromeDriverService.CreateDefaultService(Environment.CurrentDirectory);
			service.HideCommandPromptWindow = true;
			_driver = new ChromeDriver(service, options);
			_driver.Url = "https://clients3.google.com/cast/chromecast/home";
			_driver.Navigate();
		}

		static void ExtractChromeDriver()
		{
			if (!File.Exists(_chromeDriverExe))
			{
				File.WriteAllBytes(_chromeDriverExe, Resources.chromedriver);
			}
		}
		static void ChangeWallPaper()
		{
			try
			{
				var pathElement = _driver.FindElementById("picture-background");
				var url = pathElement.GetAttribute("src");
				if (_currentImageUrl != url)
				{
					_currentImageUrl = url;
					using (WebClient client = new WebClient())
					{
						IActiveDesktop iad = shlobj.GetActiveDesktop();
						client.DownloadFile(new Uri(url), "wallpaper");
						iad.SetWallpaper(Path.Combine(Environment.CurrentDirectory, "wallpaper"), 0);
						WALLPAPEROPT wopt = new WALLPAPEROPT();
						iad.GetWallpaperOptions(ref wopt, 0);
						wopt.dwStyle = WallPaperStyle.WPSTYLE_STRETCH;
						wopt.dwSize = WALLPAPEROPT.SizeOf;
						iad.SetWallpaperOptions(ref wopt, 0);
						iad.ApplyChanges(AD_Apply.ALL | AD_Apply.FORCE | AD_Apply.BUFFERED_REFRESH);
					}
				}
			}
			catch
			{
				//driver died most likely, it will get restarted
			}
		}

		private static bool IsDriverValid()
		{
			try
			{
				var handles = _driver.WindowHandles;
				return true;
			}
			catch
			{
				return false;
			}
		}

		public class MyCustomApplicationContext : ApplicationContext
		{
			private readonly NotifyIcon _trayIcon;

			public MyCustomApplicationContext()
			{
				_trayIcon = new NotifyIcon()
				{
					Icon = ChromecastWallpapers.Properties.Resources.Icon,
					ContextMenu = new ContextMenu(new MenuItem[] {
						new MenuItem("Exit", Exit)
					}),
					Visible = true
				};
			}

			private void Exit(object sender, EventArgs e)
			{
				Timer.Stop();
				Timer.Dispose();
				_driver.Close();
				_driver.Dispose();
				_trayIcon.Visible = false;
				Application.Exit();
			}
		}
	}
	
}
