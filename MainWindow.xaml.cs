using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Win32;
using SHDocVw;

namespace ExternalWebView
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<string, string> argsDic = new Dictionary<string, string>();
        private string[] keys = new string[] { "url", "posx", "posy", "width", "height", "ie" };

        private IWebBrowser2 iwBrowser = null;

        [ComImport, Guid("6d5140c1-7436-11ce-8034-00aa006009fa"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IServiceProvider
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object QueryService(ref Guid guidService, ref Guid riid);
        }


        public MainWindow()
        {
            //引数の確認
            string[] args = System.Environment.GetCommandLineArgs();
            if (args.Length <= 1)
            {
                MessageBox.Show(
                    "Usage: [ARGS] --url, --posx, --posy, --width, --height, --ie\n"
                    + "\n"
                    + "ex) --url=http://www.google.com --posx=0 --posy=0 --width=640 --height=480 --ie=10");
                return;
            }

            //引数をハッシュに変換
            for (int i = 1; i < args.Length; i++)
            {
                foreach (string k in keys)
                {
                    if (args[i].Contains(k))
                    {
                        string v = args[i].Split(new string[] { "=" }, StringSplitOptions.None)[1];
                        argsDic.Add(k, v);

                        break;
                    }
                }
            }

            //内部IEバージョンを設定する
            if(argsDic.ContainsKey("ie"))
            {
                string ie = argsDic["ie"];
                int ver;
                switch(ie)
                {
                    case "7": ver = 7000; break;
                    case "8": ver = 8888; break;
                    case "9": ver = 9999; break;
                    case "10": ver = 10001; break;
                    case "11": ver = 11001; break;
                    default: ver = 10001; break;
                }

                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true);
                key.SetValue("ExternalWebView.exe", ver, RegistryValueKind.DWord);
            }
            else
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION");
                key.DeleteValue("ExternalWebView.exe");
            }
            
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //windowの非表示
            this.Visibility = System.Windows.Visibility.Hidden;

            //windowサイズ設定
            double x = 0;
            if (argsDic.ContainsKey("posx"))
                double.TryParse(argsDic["posx"], out x);
            double y = 0;
            if (argsDic.ContainsKey("posy"))
                double.TryParse(argsDic["posx"], out y);
            double w = 0;
            if (argsDic.ContainsKey("width"))
                double.TryParse(argsDic["width"], out w);
            double h = 0;
            if (argsDic.ContainsKey("height"))
                double.TryParse(argsDic["height"], out h);
            this.Top = x;
            this.Left = y;
            this.Width = w;
            this.Height = h;
            
            //新規windowの場合に元のwindowで上書き遷移させる
            this.webBrowser.LoadCompleted += webBrowser_LoadCompleted;

            //urlの遷移
            if (argsDic.ContainsKey("url"))
            {
                string url = argsDic["url"];
                if(!url.Contains("http://") && !url.Contains("https://"))
                    url = System.IO.Path.Combine(Directory.GetCurrentDirectory(), url);

                this.webBrowser.Navigate(url);
            }
        }

        private void webBrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
            if (iwBrowser == null)
            {
                if (this.webBrowser.Document == null)
                    return;

                IServiceProvider serviceProvider = (IServiceProvider)this.webBrowser.Document;
                Guid serviceGuid = new Guid("0002DF05-0000-0000-C000-000000000046");
                Guid iid = typeof(IWebBrowser2).GUID;
                iwBrowser = (IWebBrowser2)serviceProvider.QueryService(ref serviceGuid, ref iid);

                ((DWebBrowserEvents_Event)iwBrowser).NewWindow += new DWebBrowserEvents_NewWindowEventHandler(iwBrowser_NewWindow);
            }
        }

        private void iwBrowser_NewWindow(string url, int flags, string targetFrameName, ref object postData, string headers, ref  bool processed)
        {
            //上書き遷移
            processed = true;
            this.webBrowser.Navigate(url);
        }

        private void webBrowser_Navigated(object sender, NavigationEventArgs e)
        {
            //JSエラーを無効化
            FieldInfo fld = typeof(System.Windows.Controls.WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fld != null)
            {
                object obj = fld.GetValue(this.webBrowser);
                if (obj != null)
                    obj.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, obj, new object[]{ true });
            }

            //windowの表示
            this.Visibility = System.Windows.Visibility.Visible;
        }
    }
}
