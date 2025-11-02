using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Windows.Forms;
//这是实时网页窗口
namespace PowerPointAddIn
{
    public partial class WebBrowserForm : Form
    {
        public WebBrowserForm()
        {
            this.AutoScaleMode = AutoScaleMode.Dpi;
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            InitializeAsync();
        }
        // 导入Windows API函数
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOPMOST = 0x00000008;
        private const int GWL_STYLE = -16;
        private const int WS_CHILD = 0x40000000;

        public void SetAsChildWindow(IntPtr parentHandle)
        {
            // 设置父窗口
            SetParent(this.Handle, parentHandle);

            // 设置窗口样式为子窗口
            int style = GetWindowLong(this.Handle, GWL_STYLE);
            SetWindowLong(this.Handle, GWL_STYLE, style | WS_CHILD);

            // 确保窗口不激活
            style = GetWindowLong(this.Handle, GWL_EXSTYLE);
            SetWindowLong(this.Handle, GWL_EXSTYLE, (style | WS_EX_TOPMOST) & ~WS_EX_NOACTIVATE);

            // 确保窗体出现在正确的显示器上
            this.StartPosition = FormStartPosition.Manual;
        }
        private async void InitializeAsync()
        {
            try
            {
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointAddIn",
                    "WebView2Data");

                // 确保目录存在并设置权限
                Directory.CreateDirectory(appDataPath);

                // 设置目录权限为完全控制
                var directoryInfo = new DirectoryInfo(appDataPath);
                var directorySecurity = directoryInfo.GetAccessControl();
                directorySecurity.AddAccessRule(
                    new FileSystemAccessRule(
                        Environment.UserName,
                        FileSystemRights.FullControl,
                        InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                        PropagationFlags.None,
                        AccessControlType.Allow));
                directoryInfo.SetAccessControl(directorySecurity);

                var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                    userDataFolder: appDataPath,
                    browserExecutableFolder: null,
                    options: new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions("--disable-web-security"));

                await webView21.EnsureCoreWebView2Async(env);
                webView21.CoreWebView2.NavigationCompleted += webView21_NavigationCompleted;
                webView21.CoreWebView2.NewWindowRequested += CoreWebView2_NewWindowRequested;
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"WebView2初始化失败: {ex.Message}\nStackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    //    MessageBox.Show($"Inner Exception: {ex.InnerException.Message}");
                }
            }
        }
        private void CoreWebView2_NewWindowRequested(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs e)
        {
            e.Handled = true;
            webView21.CoreWebView2.Navigate(e.Uri);
        }
        private void webView21_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            // 导航完成后的处理
        }
        public void AddWebView(string url, int width, int height)
        {
            try
            {
                if (webView21.CoreWebView2 != null)
                {
                    webView21.Source = new Uri(url);
                }
                else
                {
                    webView21.CoreWebView2InitializationCompleted += (s, e) =>
                    {
                        if (e.IsSuccess)
                        {
                            webView21.Source = new Uri(url);
                        }
                    };
                }

                webView21.Width = width;
                webView21.Height = height;
                this.ClientSize = new System.Drawing.Size(width, height);
                webView21.Dock = DockStyle.Fill;
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"加载网页失败: {ex.Message}");
            }
        }
    }
}