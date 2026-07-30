using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Windows.Forms;

namespace PowerPointAddIn
{
    public partial class WebBrowserForm : Form
    {
        public WebBrowserForm()
        {
            InitializeComponent();
            // 关闭自动缩放与自动尺寸：
            // 我们会用 SetWindowPos 以像素精确设置窗口位置和大小，
            // AutoScaleMode/AutoSize 会在不同 DPI 下二次改变尺寸，导致无法精确覆盖占位符。
            this.AutoScaleMode = AutoScaleMode.None;
            this.AutoSize = false;
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.FormClosing += WebBrowserForm_FormClosing;
            // 注意：不要在这里调用 SetProcessDPIAware()。
            // 插件是 in-process 加载到 PowerPoint 的，强制设置“系统级 DPI 感知”
            // 会覆盖 PowerPoint 自身的“每监视器 DPI 感知 v2(PMv2)”，导致在 DPI
            // 与主屏不同的副屏上，子窗口与 WebView2 内容尺寸/清晰度异常。
            // 让 PowerPoint 的 PMv2 保持生效，子窗口会继承放映窗口所在屏幕的 DPI。
            InitializeAsync();
        }
        // 导入Windows API函数
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOPMOST = 0x00000008;
        private const int GWL_STYLE = -16;
        private const int GWL_HWNDPARENT = -8;
        private const int WS_CHILD = 0x40000000;
        private void WebBrowserForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 清理WebView2资源
            try
            {
                if (webView21 != null && webView21.CoreWebView2 != null)
                {
                    webView21.CoreWebView2.Stop();
                }
                webView21.Dispose();
            }
            catch
            {
                // 忽略清理错误
            }
        }
        public string CurrentUrl { get; private set; }

        // 记录当前覆盖层所依附的放映窗口句柄，便于翻页/定时刷新时重新定位，
        // 以及检测该窗口是否仍存在（演讲者视图关闭或窗口重建时句柄会失效）。
        public IntPtr OwnerHwnd { get; set; }

        // 上一次“有效定位”的结果，用于在换页过渡、窗口几何暂不可得时，
        // 仍能停留在正确位置并可见（避免退回屏幕尺寸/1px 而闪退到屏幕外导致“看不见”）。
        public int LastGoodLeft, LastGoodTop, LastGoodWidth, LastGoodHeight;
        public bool HasLastGood = false;

        public void NavigateToUrl(string url)
        {
            try
            {
                if (webView21.CoreWebView2 != null)
                {
                    webView21.Source = new Uri(url);
                    CurrentUrl = url;
                }
                else
                {
                    webView21.CoreWebView2InitializationCompleted += (s, e) =>
                    {
                        if (e.IsSuccess)
                        {
                            webView21.Source = new Uri(url);
                            CurrentUrl = url;
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导航失败: {ex.Message}");
            }
        }

        public void AddWebView(string url, int width, int height)
        {
            try
            {
                CurrentUrl = url;

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
                MessageBox.Show($"加载网页失败: {ex.Message}");
            }
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);

            if (this.Visible && webView21 != null && webView21.CoreWebView2 != null)
            {
            }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    if (webView21 != null)
                    {
                        webView21.Dispose();
                    }
                }
                catch
                {
                    // 忽略清理错误
                }
            }
            base.Dispose(disposing);
        }
        public void SetAsOverlayWindow(IntPtr parentHandle)
        {
            // 作为“拥有者窗口(owner)”附加到放映窗口，而不是 WS_CHILD 子窗口：
            // 拥有者窗口始终位于放映窗口之上、并随其移动/关闭；
            // 而 WS_CHILD 子窗口在跨“不同 DPI 的屏幕”时，其坐标会被父窗口的
            // DPI 上下文二次缩放，正是多显示器下网页“水平偏左”的根因。
            int style = GetWindowLong(this.Handle, GWL_STYLE);
            SetWindowLong(this.Handle, GWL_STYLE, style & ~WS_CHILD);

            // 设为放映窗口的拥有者（owner），保持置顶跟随
            SetWindowLongPtr(this.Handle, GWL_HWNDPARENT, parentHandle);
            OwnerHwnd = parentHandle;

            // 置顶 + 不激活
            int exStyle = GetWindowLong(this.Handle, GWL_EXSTYLE);
            SetWindowLong(this.Handle, GWL_EXSTYLE, exStyle | WS_EX_TOPMOST | WS_EX_NOACTIVATE);

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
        private void CoreWebView2_NewWindowRequested(object sender,Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs e)
        {
            e.Handled = true;
            webView21.CoreWebView2.Navigate(e.Uri);
        }
        private void webView21_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            // 导航完成后的处理
        }
        
    }
}