using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.WebSockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn
{
    public partial class ThisAddIn
    {
        private Slide _currentSlide = null;
        private System.Windows.Forms.Timer _slideMonitorTimer;
        private static WinEventDelegate _winEventDelegate;
        private static readonly IntPtr HWND_TOP = IntPtr.Zero;
        private IntPtr _winEventHook = IntPtr.Zero;
        private Dictionary<Slide, WebBrowserForm> _activeWebViews = new Dictionary<Slide, WebBrowserForm>();
        // 演讲者视图等"额外放映窗口"各自的覆盖层（与全屏窗口分开存储，
        // 避免翻页后覆盖层被重新挂到全屏窗口而丢失演讲者视图上的显示）。
        private Dictionary<Slide, WebBrowserForm> _presenterViewWebViews = new Dictionary<Slide, WebBrowserForm>();
        // 防抖计数器：演讲者视图窗口在换页过渡瞬间会瞬态地"消失"(Count 降到 1)，
        // 仅当持续多拍都缺窗口时才判定其真的关闭，避免瞬时缺窗口把覆盖层误隐藏。
        private int _presenterMissingTicks = 0;
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        private bool _isInSlideShow = false;
        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;//添加网页的主面板
        private WebViewPanel _webViewPanel;//添加网页的主面板
        private Microsoft.Office.Tools.CustomTaskPane _qrCodePane;//二维码面板
        private QRCodePanel _qrCodePanel;//二维码面板
        private Microsoft.Office.Tools.CustomTaskPane _symbolPane;//符号面板
        private SymbolPanel _symbolPanel;//符号面板

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);

        private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        // 获取窗口所在监视器（用于按"全屏所在屏幕"的 DPI/尺寸计算）
        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;

        // 获取指定监视器的有效 DPI（每监视器 DPI 感知）
        [DllImport("user32.dll")]
        private static extern bool GetDpiForMonitor(IntPtr hMonitor, int dpiType, out uint dpiX, out uint dpiY);
        private const int MDT_EFFECTIVE_DPI = 0;

        // 获取某个窗口自身生效的 DPI
        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        private System.Drawing.Rectangle GetSlideShowScreenBounds()
        {
            if (Application.SlideShowWindows.Count == 0)
                return Screen.PrimaryScreen.Bounds;

            try
            {
                IntPtr slideShowHwnd = (IntPtr)Application.SlideShowWindows[1].HWND;
                GetWindowRect(slideShowHwnd, out RECT windowRect);

                // 将窗口矩形转换为Rectangle
                var windowBounds = new System.Drawing.Rectangle(
                    windowRect.Left, windowRect.Top,
                    windowRect.Right - windowRect.Left,
                    windowRect.Bottom - windowRect.Top);

                // 找到包含大部分窗口的显示器
                Screen targetScreen = Screen.FromRectangle(windowBounds);

                // 如果窗口跨多个显示器，选择包含窗口中心点的显示器
                var centerPoint = new System.Drawing.Point(
                    windowRect.Left + (windowRect.Right - windowRect.Left) / 2,
                    windowRect.Top + (windowRect.Bottom - windowRect.Top) / 2);

                targetScreen = Screen.FromPoint(centerPoint);

                return targetScreen.Bounds;
            }
            catch
            {
                return Screen.PrimaryScreen.Bounds;
            }
        }
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            WriteTrace("==========================================================");
            WriteTrace($"=== 新会话开始 {BuildMarker} ===");
            WriteTrace($"程序集位置={System.Reflection.Assembly.GetExecutingAssembly().Location}");
            WriteTrace($"日志路径={LogFilePath}");
            WriteTrace($"OS={Environment.OSVersion} 64位OS={Environment.Is64BitOperatingSystem} 64位进程={Environment.Is64BitProcess}");
            WriteTrace($"用户={Environment.UserDomainName}\\{Environment.UserName} CPU核心={Environment.ProcessorCount}");
            try
            {
                var wv2Asm = typeof(Microsoft.Web.WebView2.Core.CoreWebView2Environment).Assembly;
                WriteTrace($"WebView2 SDK 程序集版本={wv2Asm.GetName().Version} 位置={wv2Asm.Location}");
                string runtimeVer = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString();
                WriteTrace($"WebView2 运行时版本={runtimeVer}");
            }
            catch (Exception ex)
            {
                WriteTrace($"读取WebView2版本失败(可能未安装运行时): {ex.Message}");
            }
            WriteTrace("==========================================================");
            // 首次使用本版本时弹窗确认，便于确认 PowerPoint 加载的是最新 DLL。
            // 之后同一版本不再重复弹出（记录在日志文件的版本标记里）。
            string markerFile = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PPT-Addin", "marker.txt");
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(markerFile));
                string prev = File.Exists(markerFile) ? File.ReadAllText(markerFile).Trim() : null;
                if (prev != BuildMarker)
                {
                    File.WriteAllText(markerFile, BuildMarker);
                    MessageBox.Show(
                        $"网页插件版本：{BuildMarker}\n程序集位置：\n{System.Reflection.Assembly.GetExecutingAssembly().Location}\n\n" +
                        "看到这个提示说明 PowerPoint 加载的是最新代码。\n若未出现此提示，说明仍在用旧 DLL。",
                        "PPT网页插件 - 版本确认", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch { }

            _webViewPanel = new WebViewPanel();
            _webViewPanel.AddWebViewRequested += AddWebViewToSlide;

            _taskPane = this.CustomTaskPanes.Add(_webViewPanel, "网页浏览器");
            _taskPane.VisibleChanged += TaskPane_VisibleChanged;
            _taskPane.Visible = true;

            Application.SlideShowBegin += Application_SlideShowBegin;
            Application.SlideShowNextSlide += Application_SlideShowNextSlide;
            Application.SlideShowEnd += Application_SlideShowEnd;
            Application.PresentationOpen += Application_PresentationOpen;

            // 将委托保存到静态变量
            _winEventDelegate = new WinEventDelegate(WinEventProc);

            // 使用保存的委托
            _winEventHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
            try
            {
                var activeWindow = Application.ActiveWindow;
                if (activeWindow != null)
                {
                    activeWindow.Activate();
                }
            }
            catch { }

        }
        // 窗口焦点变化回调
        private void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            try
            {
                if (_isInSlideShow && Application.SlideShowWindows.Count > 0)
                {
                    var slideShowHwnd = (IntPtr)Application.SlideShowWindows[1].HWND;

                    if (hwnd == slideShowHwnd)
                    {
                        // 使用BeginInvoke确保在主线程执行
                        System.Windows.Forms.Application.Idle += OnIdleShowWebView;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WinEventProc error: {ex.Message}");
            }
        }
        private void InitializeSlideMonitor()
        {
            _slideMonitorTimer = new System.Windows.Forms.Timer();
            _slideMonitorTimer.Interval = 200; // 每200毫秒检查一次
            _slideMonitorTimer.Tick += (sender, e) => CheckCurrentSlide();
            _slideMonitorTimer.Start();
        }
        private void CheckCurrentSlide()
        {
            try
            {
                if (!_isInSlideShow || Application.SlideShowWindows.Count == 0)
                {
                    // 不在放映状态：仅隐藏所有网页覆盖层（保留 WebView2 实例与已加载页面），
                    // 以便退出放映模式但 PowerPoint 仍运行时，重新进入放映可即时复用、无需重新加载。
                    // 真正释放只在 PowerPoint 关闭（ThisAddIn_Shutdown）时进行。
                    HideAllWebViews();
                    return;
                }

                var view = Application.SlideShowWindows[1].View;
                var currentState = GetSlideShowState(view);

                // 如果是黑屏/白屏/结束状态，隐藏所有网页
                if (currentState != SlideShowState.Normal)
                {
                    HideAllWebViews();
                    return;
                }

                var currentSlide = view.Slide;

                // 如果幻灯片没有变化，仍要持续重新对齐（放映窗口可能被拖到
                // 另一块 DPI/尺寸不同的屏幕，或缩放倍率发生变化）。
                if (_currentSlide != null && _currentSlide.SlideID == currentSlide.SlideID)
                {
                    RepositionWebViewForSlide(currentSlide);
                    return;
                }

                _currentSlide = currentSlide;

                // 检查当前幻灯片是否有网页占位符
                bool hasWebView = false;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in currentSlide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_") && !string.IsNullOrEmpty(shape.Tags["WebViewUrl"]))
                    {
                        hasWebView = true;
                        ShowWebViewForCurrentSlide();
                        break;
                    }
                }

                // 如果没有网页占位符，隐藏所有网页（包括当前幻灯片的）
                if (!hasWebView)
                {
                    DisposeWebViewForSlide(currentSlide);
                    HideAllWebViews();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"检查幻灯片时出错: {ex.Message}");
            }
        }
        private SlideShowState GetSlideShowState(SlideShowView view)
        {
            try
            {
                if (view.State == PpSlideShowState.ppSlideShowBlackScreen)
                    return SlideShowState.BlackScreen;
                if (view.State == PpSlideShowState.ppSlideShowWhiteScreen)
                    return SlideShowState.WhiteScreen;
                if (view.State == PpSlideShowState.ppSlideShowDone)
                    return SlideShowState.Ended;

                return SlideShowState.Normal;
            }
            catch
            {
                return SlideShowState.Ended;
            }
        }
        private void HideAllWebViews()
        {
            HideAllInDict(_activeWebViews);
            HideAllInDict(_presenterViewWebViews);
        }
        private void HideAllInDict(Dictionary<Slide, WebBrowserForm> dict)
        {
            List<Slide> slidesToRemove = new List<Slide>();

            foreach (var kvp in dict)
            {
                try
                {
                    var form = kvp.Value;
                    if (!form.IsDisposed)
                    {
                        if (form.InvokeRequired)
                        {
                            form.Invoke(new Action(() => form.Hide()));
                        }
                        else
                        {
                            form.Hide();
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"隐藏WebView时出错: {ex.Message}");
                    slidesToRemove.Add(kvp.Key);
                }
            }

            // 清理已处理的条目
            foreach (var slide in slidesToRemove)
            {
                dict.Remove(slide);
            }
        }
        /// <summary>
        /// 仅隐藏除指定幻灯片以外，所有幻灯片在"全屏 + 演讲者视图"上的 WebView 覆盖层。
        /// 与 DisposeAllWebViewsExcept 不同，这里【不释放】WebView2 实例与已加载页面，
        /// 以满足"切换页面 / 退出放映时保留已加载页面、仅隐藏，下次进入瞬时复用"的要求。
        /// </summary>
        private void HideAllWebViewsExcept(Slide exceptSlide)
        {
            HideAllExceptInDict(_activeWebViews, exceptSlide);
            HideAllExceptInDict(_presenterViewWebViews, exceptSlide);
        }
        private void HideAllExceptInDict(Dictionary<Slide, WebBrowserForm> dict, Slide exceptSlide)
        {
            foreach (var kvp in dict)
            {
                try
                {
                    if (kvp.Key.SlideID == exceptSlide.SlideID) continue; // 当前幻灯片：保留，由后续显示逻辑负责呈现
                    var form = kvp.Value;
                    if (!form.IsDisposed && form.Visible)
                    {
                        if (form.InvokeRequired)
                        {
                            form.Invoke(new Action(() => form.Hide()));
                        }
                        else
                        {
                            form.Hide();
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"隐藏WebView时出错: {ex.Message}");
                }
            }
        }
        private void OnIdleShowWebView(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Idle -= OnIdleShowWebView;
            try
            {
                ShowWebViewForCurrentSlide();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnIdleShowWebView error: {ex.Message}");
            }
        }
        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            // 如果需要，可以在这里更新Ribbon按钮状态
            if (ribbon != null)
            {
                // 刷新Ribbon UI
            }
        }
        private Ribbon1 ribbon;
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon1();
            return ribbon;
        }
        public void ToggleTaskPane()
        {
            _taskPane.Visible = !_taskPane.Visible;
        }
        private void Application_PresentationOpen(Presentation Pres)
        {
            // 只存储形状信息，不创建WebView
            foreach (Slide slide in Pres.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_"))
                    {
                        // 确保形状有必要的标签信息
                        if (string.IsNullOrEmpty(shape.Tags["WebViewUrl"]))
                        {
                            shape.Tags.Add("WebViewUrl", "");
                        }
                    }
                }
            }
        }
        private void Application_SlideShowBegin(SlideShowWindow Wn)
        {
            _isInSlideShow = true;
            _presenterMissingTicks = 0; // 重置防抖计数器
            InitializeSlideMonitor(); // 开始监控幻灯片
            // 幻灯片放映开始时，检查当前幻灯片是否有网页控件
            var currentSlide = Wn.View.Slide;
            ShowWebViewForCurrentSlide();
        }
        private void AddWebViewToSlide(string url, int width, int height)
        {
            try
            {
                var slide = Application.ActiveWindow.View.Slide;
                var shapes = slide.Shapes;

                var shape = shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    0, 0, width, height);

                string shapeName = "WebViewContainer_" + Guid.NewGuid().ToString();
                shape.Name = shapeName;

                // 将URL和尺寸信息存储在形状的标签中
                shape.Tags.Add("WebViewUrl", url);
                shape.Tags.Add("WebViewWidth", width.ToString());
                shape.Tags.Add("WebViewHeight", height.ToString());

                // 添加灰色背景和边框
                shape.Fill.ForeColor.RGB = System.Drawing.Color.LightGray.ToArgb() & 0xFFFFFF;
                shape.Fill.Visible = MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = System.Drawing.Color.DarkGray.ToArgb() & 0xFFFFFF;
                shape.Line.Visible = false;//想要边框的自己修改这行为MsoTriState.msoTrue，然后上一行自己选颜色
                shape.Line.Weight = 1.00f;

                // 只在幻灯片放映时创建WebView
                if (_isInSlideShow && Application.SlideShowWindows.Count > 0)
                {
                    // 先移除同 SlideID 的旧记录，确保同一张幻灯片在字典中至多一条
                    RemoveBySlideId(_activeWebViews, slide.SlideID);
                    var webForm = new WebBrowserForm();
                    webForm.AddWebView(url, width, height);
                    _activeWebViews[slide] = webForm;
                    ShowWebViewForCurrentSlide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加网页失败: {ex.Message}");
            }
        }
        private void Application_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (_isInSlideShow)
            {
                ShowWebViewForCurrentSlide();
            }
        }
        private enum SlideShowState
        {
            Normal,
            BlackScreen,
            WhiteScreen,
            Ended
        }

        /// <summary>
        /// 相对路径存储前缀。带此前缀的 WebViewUrl 表示"相对于 PPT 所在目录"的路径，
        /// 在显示时再实时解析为绝对 file URI，从而支持"PPT 与网页一起整体移动"。
        /// </summary>
        public const string RelativePathPrefix = "relative://";

        /// <summary>
        /// 代码版标记。每次改动相对路径逻辑请递增，用于确认 PowerPoint 加载的是最新 DLL。
        /// </summary>
        private const string BuildMarker = "RELATIVE-FIX-v6-IPV4";

        /// <summary>
        /// 诊断日志：记录 WebView2 环境、导航、资源加载、进程崩溃、JS 控制台错误等，
        /// 用于排查"网页资源加载失败"。日志位于 %LOCALAPPDATA%\PPT-Addin\webview.log
        /// </summary>
        public static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PPT-Addin", "webview.log");

        private static readonly object _logLock = new object();

        public static void WriteTrace(string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(LogFilePath);
                Directory.CreateDirectory(dir);

                // 超过 5MB 则归档，避免日志无限增长影响排查
                try
                {
                    var fi = new FileInfo(LogFilePath);
                    if (fi.Exists && fi.Length > 5 * 1024 * 1024)
                    {
                        string oldPath = Path.Combine(dir, "webview.old.log");
                        if (File.Exists(oldPath)) File.Delete(oldPath);
                        File.Move(LogFilePath, oldPath);
                    }
                }
                catch { }

                lock (_logLock)
                {
                    File.AppendAllText(LogFilePath,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}\r\n");
                }
            }
            catch { }
        }

        /// <summary>
        /// 将 HTML 文件的 file:// 绝对地址转换为"相对于当前 PPT 所在目录"的存储串(带前缀)。
        /// 按需求：仅在 PPT 同目录下查找，因此只保留文件名（忽略任何子目录结构）。
        /// 仅当 PPT 已保存且网页位于 PPT 目录内时返回有效串；否则返回 null（调用方回退为绝对路径）。
        /// </summary>
        public string MakeRelativeWebViewUrl(string absoluteFileUrl)
        {
            try
            {
                string htmlPath = new Uri(absoluteFileUrl).LocalPath;
                string pptDir = Application.ActivePresentation.Path;
                if (string.IsNullOrEmpty(pptDir))
                {
                    WriteTrace($"[转相对] 失败：PPT尚未保存，无所在目录。文件={htmlPath}");
                    return null;
                }
                if (!htmlPath.StartsWith(pptDir, StringComparison.OrdinalIgnoreCase))
                {
                    WriteTrace($"[转相对] 失败：网页不在PPT目录下。PPT目录={pptDir} 网页={htmlPath}");
                    return null;
                }
                // 仅取文件名（忽略子目录），后续解析只在 PPT 同目录查找
                string fileName = Path.GetFileName(htmlPath);
                if (string.IsNullOrEmpty(fileName)) return null;
                WriteTrace($"[转相对] 成功：{absoluteFileUrl} -> {RelativePathPrefix}{fileName} （PPT目录={pptDir}）");
                return RelativePathPrefix + fileName;
            }
            catch (Exception ex)
            {
                WriteTrace($"[转相对] 异常：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 解析存储的网页地址：若为 relative:// 前缀，则只在"当前 PPT 所在目录"下按文件名查找，
        /// 解析为绝对 file URI；否则原样返回（绝对路径或 http(s) 等）。
        /// PPT 与网页一起移动后，此处会按新目录定位（仅限同目录）。
        /// </summary>
        public string ResolveWebViewUrl(string storedUrl)
        {
            if (string.IsNullOrEmpty(storedUrl)) return storedUrl;

            try
            {
                string pptDir = null;
                try { pptDir = Application.ActivePresentation.Path; } catch { }
                if (string.IsNullOrEmpty(pptDir)) pptDir = null;

                // 1) relative:// 前缀：只在 PPT 同目录按文件名查找
                if (storedUrl.StartsWith(RelativePathPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    string fileName = Path.GetFileName(
                        storedUrl.Substring(RelativePathPrefix.Length).Replace('/', '\\'));
                    string fullPath = pptDir != null
                        ? Path.Combine(pptDir, fileName)
                        : Path.Combine(Directory.GetCurrentDirectory(), fileName);
                    var uri = new Uri(fullPath).AbsoluteUri;
                    WriteTrace($"[解析相对] PPT目录={pptDir} 文件名={fileName} -> {fullPath} 存在={File.Exists(fullPath)}");
                    return uri;
                }

                // 2) 历史 PPT 存的是绝对 file:// 地址：若该地址的文件已不存在（说明 PPT 曾被移动），
                //    自动回退到"PPT 同目录 + 同名文件"，这样老文件无需重新添加也能正常放映。
                if (pptDir != null && storedUrl.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        string localPath = new Uri(storedUrl).LocalPath;
                        if (!string.IsNullOrEmpty(localPath) && !File.Exists(localPath))
                        {
                            string candidate = Path.Combine(pptDir, Path.GetFileName(localPath));
                            WriteTrace($"[历史迁移] 原绝对路径文件不存在: {localPath}；PPT同目录候选={candidate} 存在={File.Exists(candidate)}");
                            if (File.Exists(candidate))
                            {
                                WriteTrace($"[历史迁移] 已改用PPT同目录同名文件: {candidate}");
                                return new Uri(candidate).AbsoluteUri;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return storedUrl;
        }

        private void ShowWebViewForCurrentSlide()
        {
            if (!_isInSlideShow || Application.SlideShowWindows.Count == 0) return;

            try
            {
                var slideShowWindow = Application.SlideShowWindows[1];
                var currentSlide = slideShowWindow.View.Slide;

                // 获取放映窗口所在的屏幕信息
                var screenBounds = GetSlideShowScreenBounds();
                var screen = Screen.FromRectangle(screenBounds);

                // 仅隐藏其它幻灯片的网页覆盖层，【不释放】其 WebView2 实例与已加载页面。
                // 这样在同一次放映内切换页面、或退出放映模式后再次进入时，已加载的页面保持驻留，
                // 重新显示即可瞬时复用、无需重新加载、不浪费网络开销。
                // 所有 WebView2 实例只在 PowerPoint 关闭（ThisAddIn_Shutdown）时才统一释放。
                HideAllWebViewsExcept(currentSlide);

                // 检查当前幻灯片是否有网页控件（只取第一个 WebViewContainer_ 占位符）
                Microsoft.Office.Interop.PowerPoint.Shape webShape = null;
                string webUrl = null;
                int webWidth = 0, webHeight = 0;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in currentSlide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_"))
                    {
                        string url = shape.Tags["WebViewUrl"];
                        if (!string.IsNullOrEmpty(url))
                        {
                            webShape = shape;
                            webUrl = ResolveWebViewUrl(url);
                            webWidth = int.Parse(shape.Tags["WebViewWidth"]);
                            webHeight = int.Parse(shape.Tags["WebViewHeight"]);
                            break;
                        }
                    }
                }

                if (webShape != null)
                {
                    // 为"每一个放映窗口"各创建一个覆盖层：
                    // 处于放映状态时窗口可能不止一个（例如演讲者视图 + 全屏观众窗口），
                    // 每个 SlideShowWindow 的客户区恰好就是"当前幻灯片"的显示区域，
                    // 因此对每个窗口都用同一套"幻灯片磅值 -> 物理像素"的映射即可精确覆盖。
                    // 这样无论是全屏还是演讲者视图，翻页后都会随页面切换刷新并显示实时网页。
                    int windowCount = Application.SlideShowWindows.Count;
                    for (int i = 1; i <= windowCount; i++)
                    {
                        // 每个窗口独立 try：某个窗口(如演讲者视图)在过渡瞬间取句柄/定位
                        // 抛异常时，不影响其它窗口，且定时器每 200ms 还会自动重试。
                        try
                        {
                            var winHwnd = (IntPtr)Application.SlideShowWindows[i].HWND;
                            // 第一个窗口作为主窗口(全屏)，其余窗口作为演讲者视图等额外窗口
                            bool isPresenter = (i != 1);

                            var form = GetOrCreateWebViewForSlide(currentSlide, webUrl, webWidth, webHeight, isPresenter);

                            // 设置为该放映窗口的拥有者窗口（顶层覆盖，跨屏不偏移）
                            if (form.InvokeRequired)
                            {
                                form.Invoke(new Action(() => form.SetAsOverlayWindow(winHwnd)));
                            }
                            else
                            {
                                form.SetAsOverlayWindow(winHwnd);
                            }

                            // 计算位置和大小（按该窗口自身的客户区/所在屏幕 DPI 精确覆盖占位符）
                            CalculateAndSetWebViewPosition(form, webShape, currentSlide, screenBounds, screen, winHwnd);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"为窗口{i}创建覆盖层失败(将重试): {ex.Message}");
                        }
                    }
                }
                else
                {
                    // 没有网页占位符：清理当前幻灯片在"所有窗口"（全屏 + 演讲者视图）上的网页
                    DisposeWebViewForSlide(currentSlide);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示网页失败: {ex.Message}");
            }
        }
        /// <summary>
        /// 释放并移除除指定幻灯片以外，所有幻灯片在"全屏 + 演讲者视图"上的 WebView 覆盖层。
        ///
        /// 重要修复：原实现只 Hide 不 Dispose，导致每张放映过的幻灯片的 WebBrowserForm
        /// （各含一个 WebView2 实例）都永久驻留内存且永不释放。放映过数十张网页幻灯片后，
        /// 会累积数十个 WebView2，内存/GDI 句柄暴涨，正是"某些设备上资源加载失败"的根因。
        /// 现改为：切换到新幻灯片时，彻底释放其它幻灯片的 WebView2，使同时存在的实例数
        /// 恒定为"当前幻灯片 × 放映窗口数"（通常 1~2 个）。
        /// </summary>
        private void DisposeAllWebViewsExcept(Slide exceptSlide)
        {
            DisposeAllExceptInDict(_activeWebViews, exceptSlide);
            DisposeAllExceptInDict(_presenterViewWebViews, exceptSlide);
        }
        private void DisposeAllExceptInDict(Dictionary<Slide, WebBrowserForm> dict, Slide exceptSlide)
        {
            List<Slide> slidesToRemove = new List<Slide>();

            foreach (var kvp in dict)
            {
                if (kvp.Key.SlideID != exceptSlide.SlideID)
                    slidesToRemove.Add(kvp.Key);
            }

            foreach (var slide in slidesToRemove)
            {
                DisposeWebViewInDict(dict, slide);
            }
        }

        /// <summary>
        /// 释放全部 WebView 覆盖层并清空两个字典（用于放映结束 / 退出放映）。
        /// </summary>
        private void DisposeAllWebViews()
        {
            DisposeAllInDict(_activeWebViews);
            DisposeAllInDict(_presenterViewWebViews);
        }
        private void DisposeAllInDict(Dictionary<Slide, WebBrowserForm> dict)
        {
            var slides = dict.Keys.ToList();
            foreach (var slide in slides)
            {
                DisposeWebViewInDict(dict, slide);
            }
            dict.Clear();
        }

        /// <summary>
        /// 按"幻灯片 ID"而非 COM 对象引用在字典中查找覆盖层。
        /// 关键（WPS 兼容）：Dictionary&lt;Slide,...&gt; 的默认比较器按对象引用比较。
        /// PowerPoint 对同一张幻灯片返回同一个 COM 包装对象，引用比较可用；
        /// 但 WPS 放映时每次取 View.Slide 都可能得到"新的包装对象"，引用比较必然失配 →
        /// 每 200ms 的监控都判定"没有已有实例"而重复新建 WebView2，旧实例又不被移除，
        /// 形成死循环并不断闪白屏。故这里统一以稳定的 SlideID 作为身份依据。
        /// </summary>
        private static bool TryGetBySlideId(Dictionary<Slide, WebBrowserForm> dict, int slideId, out WebBrowserForm form)
        {
            form = null;
            foreach (var kvp in dict)
            {
                if (kvp.Key.SlideID == slideId)
                {
                    form = kvp.Value;
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 按幻灯片 ID 从字典中移除条目（同样为 WPS 的 COM 引用不稳定而设计）。
        /// </summary>
        private static bool RemoveBySlideId(Dictionary<Slide, WebBrowserForm> dict, int slideId)
        {
            Slide key = null;
            foreach (var kvp in dict)
            {
                if (kvp.Key.SlideID == slideId)
                {
                    key = kvp.Key;
                    break;
                }
            }
            return key != null && dict.Remove(key);
        }

        /// <summary>
        /// 获取或创建某一放映窗口(全屏 / 演讲者视图)对应的 WebView 覆盖层窗体，
        /// 并导航到指定 URL。isPresenter=true 时使用演讲者视图字典。
        /// 以 SlideID 判定"是否已有实例"，确保同一张幻灯片始终复用同一个 WebView2
        /// （既避免重复创建导致的死循环/白屏，也保证切回时无需重新加载）。
        /// </summary>
        private WebBrowserForm GetOrCreateWebViewForSlide(Slide slide, string url, int width, int height, bool isPresenter)
        {
            var dict = isPresenter ? _presenterViewWebViews : _activeWebViews;
            int sid = slide.SlideID;
            if (!TryGetBySlideId(dict, sid, out WebBrowserForm form) || form == null || form.IsDisposed)
            {
                // 先清掉同 SlideID 的失效残留，保证字典中每张幻灯片至多一条记录
                RemoveBySlideId(dict, sid);
                form = new WebBrowserForm();
                int capturedSid = sid;
                form.FormClosed += (s, e) =>
                {
                    RemoveBySlideId(dict, capturedSid);
                };
                form.AddWebView(url, width, height);
                dict[slide] = form;
            }
            else if (!form.IsCurrentUrl(url))
            {
                // 仅当"归一化后"的目标地址与已加载地址确实不同才重新导航；
                // 相同则直接复用已加载页面，避免重复加载（白屏闪烁 / 浪费网络）。
                form.NavigateToUrl(url);
            }
            return form;
        }

        /// <summary>
        /// 清理某张幻灯片在"所有窗口"（全屏 + 演讲者视图）上的网页覆盖层。
        /// </summary>
        private void DisposeWebViewForSlide(Slide slide)
        {
            DisposeWebViewInDict(_activeWebViews, slide);
            DisposeWebViewInDict(_presenterViewWebViews, slide);
        }
        private void DisposeWebViewInDict(Dictionary<Slide, WebBrowserForm> dict, Slide slide)
        {
            int sid = slide.SlideID;
            if (TryGetBySlideId(dict, sid, out WebBrowserForm form))
            {
                try
                {
                    if (form.InvokeRequired)
                        form.Invoke(new Action(() => { form.Hide(); form.Dispose(); }));
                    else
                    {
                        form.Hide();
                        form.Dispose();
                    }
                }
                catch { }
                // 按 SlideID 移除（WPS 下对象引用不稳定，不能用 dict.Remove(slide)）
                RemoveBySlideId(dict, sid);
            }
        }

        private void CalculateAndSetWebViewPosition(WebBrowserForm form, Microsoft.Office.Interop.PowerPoint.Shape shape,
            Slide currentSlide, System.Drawing.Rectangle screenBounds, Screen screen, IntPtr hwnd)
        {
            // 以屏幕绝对物理像素定位（窗体为放映窗口的拥有者窗口，非 WS_CHILD 子窗口）
            PositionWebView(form, shape, currentSlide, hwnd);
        }

        /// <summary>
        /// 将 WebView 窗体精确覆盖到放映窗口中占位形状的位置。
        /// 关键设计（解决多显示器"位置偏左"）：
        /// 1) 网页窗体不再作为放映窗口的 WS_CHILD 子窗口，而是其"拥有者窗口(owned window)"，
        ///    并以【屏幕绝对物理像素】定位。子窗口在跨不同 DPI 的屏幕时会被父窗口 DPI 上下文
        ///    二次缩放坐标，正是水平偏移的根因；顶层窗口会按自身所在屏幕的 DPI 自动适配，无此问题。
        /// 2) 用 ClientToScreen(hwnd, (0,0)) 取得放映窗口客户区在桌面上的绝对物理像素原点，
        ///    再叠加"幻灯片磅值 -> 物理像素"的等比居中映射，得到占位符的绝对屏幕坐标。
        /// 3) 此映射与 PowerPoint 自身的居中缩放完全一致，因此无论该屏幕的 DPI、尺寸、缩放倍率如何，
        ///    网页都能精确覆盖占位符。
        /// </summary>
        private void PositionWebView(WebBrowserForm form, Microsoft.Office.Interop.PowerPoint.Shape shape,
            Slide currentSlide, IntPtr hwnd)
        {
            try
            {
                // 幻灯片页面尺寸(磅, 1 磅 = 1/72 英寸)
                float slideWidth = currentSlide.Parent.PageSetup.SlideWidth;
                float slideHeight = currentSlide.Parent.PageSetup.SlideHeight;

                // 客户区原点在桌面上的绝对物理像素坐标。
                // 关键：ClientToScreen(hwnd, pt) 的入参 pt 必须是"客户区相对坐标"(0,0)，
                // 而不是窗口的屏幕坐标(windowRect.Left/Top)。之前把后者当入参导致副屏偏移。
                POINT clientOrigin = new POINT { X = 0, Y = 0 };
                bool gotClientOrigin = false;
                try { gotClientOrigin = ClientToScreen(hwnd, ref clientOrigin); } catch { gotClientOrigin = false; }

                // 优先取客户区尺寸(物理像素)；取不到时退回"窗口矩形"作为客户区代理。
                RECT clientRect;
                bool gotClient = GetClientRect(hwnd, out clientRect) &&
                                 (clientRect.Right - clientRect.Left) > 0 &&
                                 (clientRect.Bottom - clientRect.Top) > 0;
                RECT windowRect;
                bool gotWindow = GetWindowRect(hwnd, out windowRect);

                // 有效定位所需的"客户区尺寸/原点"：
                // 1) 客户区几何有效 → 用客户区(最精确)；
                // 2) 客户区几何取不到但有窗口矩形 → 用窗口矩形作为客户区代理(仍能映射到占位符，不会铺满窗口)；
                // 3) 两者皆无 → 无法定位，本拍放弃，等下一拍(通常 ≤200ms)再尝试。
                float effClientW = gotClient ? (clientRect.Right - clientRect.Left)
                                             : (gotWindow ? (windowRect.Right - windowRect.Left) : 0);
                float effClientH = gotClient ? (clientRect.Bottom - clientRect.Top)
                                             : (gotWindow ? (windowRect.Bottom - windowRect.Top) : 0);
                int effOriginX = gotClientOrigin ? clientOrigin.X : (gotWindow ? windowRect.Left : 0);
                int effOriginY = gotClientOrigin ? clientOrigin.Y : (gotWindow ? windowRect.Top : 0);

                // 几何是否有效：换页过渡瞬间可能取不到有效的客户区尺寸(返回 0)。
                bool geometryValid = effClientW > 0 && effClientH > 0 && slideWidth > 0 && slideHeight > 0;

                int absLeft, absTop, width, height;
                if (geometryValid)
                {
                    // 幻灯片在放映窗口中按等比缩放并居中显示(与 PowerPoint 放映方式一致)
                    float scale = Math.Min(effClientW / slideWidth, effClientH / slideHeight);
                    float slideOffsetX = (effClientW - slideWidth * scale) / 2f;
                    float slideOffsetY = (effClientH - slideHeight * scale) / 2f;

                    // 占位符号在桌面上的绝对物理像素坐标与尺寸
                    absLeft = effOriginX + (int)Math.Round(slideOffsetX + shape.Left * scale);
                    absTop = effOriginY + (int)Math.Round(slideOffsetY + shape.Top * scale);
                    width = (int)Math.Round(shape.Width * scale);
                    height = (int)Math.Round(shape.Height * scale);
                    if (width < 1) width = 1;
                    if (height < 1) height = 1;

                    // 记住"上一次有效定位"，过渡期几何取不到时仍可停留在正确位置
                    form.LastGoodLeft = absLeft;
                    form.LastGoodTop = absTop;
                    form.LastGoodWidth = width;
                    form.LastGoodHeight = height;
                    form.HasLastGood = true;
                }
                else if (form.HasLastGood)
                {
                    // 过渡瞬间几何无效：沿用上一次有效位置，保持可见且不错位
                    // （绝不用 (0,0)/1px 兜底，否则会闪退到屏幕外而"完全看不见"）。
                    absLeft = form.LastGoodLeft;
                    absTop = form.LastGoodTop;
                    width = form.LastGoodWidth;
                    height = form.LastGoodHeight;
                }
                else
                {
                    // 连窗口矩形都取不到：放弃本拍，等下一拍(通常 ≤200ms)再尝试
                    return;
                }

                // 取得"全屏所在屏幕"的有效 DPI（每监视器），用于诊断与清晰度确认
                uint monitorDpiX = 96, monitorDpiY = 96;
                try
                {
                    IntPtr hMonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
                    if (hMonitor != IntPtr.Zero)
                        GetDpiForMonitor(hMonitor, MDT_EFFECTIVE_DPI, out monitorDpiX, out monitorDpiY);
                }
                catch { }
                System.Diagnostics.Debug.WriteLine(
                    $"WebView定位: 屏幕DPI={monitorDpiX}x{monitorDpiY}, 客户区(有效)={effClientW}x{effClientH}({(gotClient ? "精确" : "窗口代理")}), 屏幕绝对位置=({absLeft},{absTop}) 尺寸={width}x{height}");

                Action apply = () =>
                {
                    if (!form.Visible)
                        form.Visible = true;
                    // 顶层拥有者窗口：用屏幕绝对物理像素定位，HWND_TOPMOST 保证置于放映窗口之上
                    SetWindowPos(form.Handle, HWND_TOPMOST, absLeft, absTop, width, height,
                        SWP_SHOWWINDOW | SWP_NOACTIVATE);
                };

                if (form.InvokeRequired)
                    form.Invoke(apply);
                else
                    apply();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"计算位置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 同一张幻灯片时，按当前放映窗口实时重新对齐 WebView。
        /// 这样即使放映窗口被拖动到另一块 DPI/尺寸不同的屏幕、或缩放倍率变化，
        /// 网页也能始终精确覆盖占位符。
        /// </summary>
        private void RepositionWebViewForSlide(Slide slide)
        {
            RepositionWebViewInDict(_activeWebViews, slide);
            RepositionWebViewInDict(_presenterViewWebViews, slide);
        }
        /// <summary>
        /// 同一张幻灯片时，按当前放映窗口实时重新对齐 WebView。
        /// 对每个窗口（全屏 + 演讲者视图）各自用其记录的窗口句柄重新定位，
        /// 这样即使放映窗口被拖动到另一块 DPI/尺寸不同的屏幕、或跨屏移动，网页也始终精确覆盖占位符。
        /// </summary>
        private void RepositionWebViewInDict(Dictionary<Slide, WebBrowserForm> dict, Slide slide)
        {
            try
            {
                if (!_isInSlideShow || Application.SlideShowWindows.Count == 0) return;

                // 找到当前幻灯片的网页占位符（含 URL/尺寸）
                Microsoft.Office.Interop.PowerPoint.Shape webShape = null;
                string webUrl = null; int webWidth = 0, webHeight = 0;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_") && !string.IsNullOrEmpty(shape.Tags["WebViewUrl"]))
                    {
                        webShape = shape;
                        webUrl = ResolveWebViewUrl(shape.Tags["WebViewUrl"]);
                        webWidth = int.Parse(shape.Tags["WebViewWidth"]);
                        webHeight = int.Parse(shape.Tags["WebViewHeight"]);
                        break;
                    }
                }
                if (webShape == null) return;

                bool isPresenter = (dict == _presenterViewWebViews);

                // 演讲者视图窗口在窗口数量不足 2 时，通常只是"换页过渡"的瞬时态，
                // 并非真的关闭。若据此立刻隐藏覆盖层，会被瞬时缺窗口误隐藏，
                // 造成"时得时不得"。因此用计数器做防抖：仅在持续多拍(约 2s)
                // 都缺窗口时才判定演讲者视图已关闭并隐藏；过渡期间的瞬态缺失一律"保持现状"。
                if (isPresenter)
                {
                    if (Application.SlideShowWindows.Count >= 2)
                    {
                        _presenterMissingTicks = 0;
                    }
                    else
                    {
                        _presenterMissingTicks++;
                        if (_presenterMissingTicks < 10)
                            return; // 瞬态缺失(换页过渡)：保持现状，下一拍窗口恢复后再对齐/显示
                        // 持续缺失(约 2s)：演讲者视图确实已关闭，隐藏其覆盖层
                        if (dict.TryGetValue(slide, out var existing) && !existing.IsDisposed)
                        {
                            if (existing.InvokeRequired) existing.Invoke(new Action(() => existing.Hide()));
                            else existing.Hide();
                        }
                        return;
                    }
                }

                int idx = isPresenter ? 2 : 1;
                if (idx > Application.SlideShowWindows.Count) idx = 1;
                IntPtr winHwnd = (IntPtr)Application.SlideShowWindows[idx].HWND;

                // 获取或创建该窗口对应的覆盖层（演讲者视图在放映中途才出现时也能补建）
                var form = GetOrCreateWebViewForSlide(slide, webUrl, webWidth, webHeight, isPresenter);

                // 若记录的窗口句柄已失效或与当前窗口不一致（演讲者视图关闭/窗口被重建），重新挂靠
                if (form.OwnerHwnd != winHwnd || !IsWindow(form.OwnerHwnd))
                {
                    form.OwnerHwnd = winHwnd;
                    if (form.InvokeRequired) form.Invoke(new Action(() => form.SetAsOverlayWindow(winHwnd)));
                    else form.SetAsOverlayWindow(winHwnd);
                }

                PositionWebView(form, webShape, slide, form.OwnerHwnd);
            }
            catch { }
        }

        // 获取DPI缩放方法
        private float GetDpiScaleForScreen(Screen screen)
        {
            try
            {
                // 方法1：使用屏幕的设备上下文
                using (var form = new Form())
                {
                    form.Location = screen.Bounds.Location;
                    form.StartPosition = FormStartPosition.Manual;
                    form.ShowInTaskbar = false;
                    form.Visible = false;
                    form.CreateControl();

                    using (Graphics g = form.CreateGraphics())
                    {
                        float dpiX = g.DpiX;
                        float scale = dpiX / 96.0f;
                        return scale;
                    }
                }
            }
            catch
            {
                // 方法2：备用方案，获取主显示器的DPI
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                {
                    if (g != null)
                    {
                        float dpiX = g.DpiX;
                        float scale = dpiX / 96.0f;
                        return scale;
                    }
                }

                // 方法3：如果前两种方法都失败，返回系统DPI缩放
                return (float)System.Windows.Forms.SystemInformation.PrimaryMonitorSize.Width /
                       Screen.PrimaryScreen.Bounds.Width;
            }
        }
        private void Application_SlideShowEnd(Presentation Pres)
        {
            _isInSlideShow = false;
            _presenterMissingTicks = 0; // 重置防抖计数器

            // 停止监控
            if (_slideMonitorTimer != null)
            {
                _slideMonitorTimer.Stop();
                _slideMonitorTimer.Dispose();
                _slideMonitorTimer = null;
            }

            ForceHideAllWebViews();
            // 关键：退出放映模式【不再释放】WebView2 实例，只隐藏覆盖层。
            // 这样在不关闭 PowerPoint 的前提下重新进入放映模式时，网页保持已加载状态、
            // 直接复用显示，无需重新加载。仅当 PowerPoint 应用程序关闭（ThisAddIn_Shutdown）时才彻底释放。
            _currentSlide = null;
        }
        private void ForceHideAllWebViews()
        {
            ForceHideInDict(_activeWebViews);
            ForceHideInDict(_presenterViewWebViews);
        }
        private void ForceHideInDict(Dictionary<Slide, WebBrowserForm> dict)
        {
            List<Slide> slidesToRemove = new List<Slide>();

            foreach (var kvp in dict)
            {
                try
                {
                    var form = kvp.Value;
                    if (!form.IsDisposed)
                    {
                        if (form.InvokeRequired)
                        {
                            form.Invoke(new Action(() =>
                            {
                                // 隐藏但不关闭，以便保留状态
                                form.Visible = false;
                            }));
                        }
                        else
                        {
                            // 隐藏但不关闭，以便保留状态
                            form.Visible = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"强制隐藏WebView时出错: {ex.Message}");
                    slidesToRemove.Add(kvp.Key);
                }
            }

            // 清理字典（仅当出现异常时）
            foreach (var slide in slidesToRemove)
            {
                dict.Remove(slide);
            }
        }
        private void CleanupWebView2Data()
        {
            // 重要：不再删除整个 WebView2 用户数据目录。
            // 该目录是共享环境的缓存/磁盘存储，整体删除会带来两个问题：
            //   1) 目录被占用时删除失败或残留半损坏状态，导致下次"资源加载失败"；
            //   2) 清空缓存后，依赖缓存的网页在离线/弱网设备上直接加载不出资源。
            // 保留缓存可显著提升加载成功率与速度；确需清理时由用户手动删除即可。
            WriteTrace("[WebView2] 退出时保留共享缓存目录(不再清空，避免资源加载失败)");
        }
        [DllImport("user32.dll")]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_winEventHook != IntPtr.Zero)
            {
                UnhookWinEvent(_winEventHook);
                _winEventHook = IntPtr.Zero;
            }

            // 清理委托引用
            _winEventDelegate = null;

            // 停止所有Web控制台活动
            WebControl.StopAll();

            // 清理WebView2数据
            CleanupWebView2Data();

            // 关闭 PowerPoint 时彻底释放所有 WebView2 实例：
            // 满足"只有关闭 PowerPoint 应用程序才卸载网页"的要求（退出放映/退出放映模式只隐藏不释放）。
            DisposeAllWebViews();
        }
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        //网页相关模块
        public void ToggleWebConsole()
        {
            WebControl.ToggleWebConsole();

            // 通知WebViewPanel更新状态
            if (_webViewPanel != null)
            {
                if (WebControl.IsConsoleEnabled())
                {
                    string url = WebControl.GetConsoleUrl();
                    _webViewPanel.UpdateConsoleStatus(true, url);
                }
                else
                {
                    _webViewPanel.UpdateConsoleStatus(false);
                }
            }
        }
        // 鼠标事件API
        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        //网页相关模块结束
        public void ToggleSymbolPanel()
        {
            if (_symbolPanel == null)
            {
                _symbolPanel = new SymbolPanel();
                _symbolPanel.AddSymbolRequested += AddSymbolToSlide;
                _symbolPane = this.CustomTaskPanes.Add(_symbolPanel, "特殊符号");
                _symbolPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                _symbolPane.Width = 300;
            }
            _symbolPane.Visible = !_symbolPane.Visible;
            // 确保面板显示时刷新内容（特别是从隐藏状态恢复时）
            if (_symbolPane.Visible)
            {
                _symbolPanel.RefreshSymbols();
            }
        }

        private void AddSymbolToSlide(string symbol, System.Drawing.Font font)
        {
            try
            {
                // 获取当前活动窗口
                var window = Application.ActiveWindow;

                // 情况1：当前正在编辑文本框
                if (window.Selection.Type == PpSelectionType.ppSelectionText)
                {
                    // 在当前位置插入符号
                    var textRange = window.Selection.TextRange;
                    var newRange = textRange.InsertAfter(symbol);

                    // 设置字体一致
                    newRange.Font.Name = font.Name;
                    newRange.Font.Size = font.Size;
                }
                // 情况2：当前没有选中任何文本元素
                else
                {
                    // 获取当前视图中的点击位置或默认位置
                    float left = 100, top = 100;

                    // 创建新文本框
                    var textBox = window.View.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,left, top, 50, 50);

                    textBox.TextFrame.TextRange.Text = symbol;
                    textBox.TextFrame.TextRange.Font.Name = font.Name;
                    textBox.TextFrame.TextRange.Font.Size = font.Size;

                    // 如果是普通视图，立即进入编辑模式
                    if (window.ViewType == PpViewType.ppViewNormal)
                    {
                        textBox.TextFrame.TextRange.Select();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加符号失败: {ex.Message}");
            }
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        private System.Drawing.Point GetCursorPosition()
        {
            POINT cursorPos;
            GetCursorPos(out cursorPos);
            return new System.Drawing.Point(cursorPos.X, cursorPos.Y);
        }
        public void ToggleQRCodePanel()
        {
            if (_qrCodePane == null)
            {
                _qrCodePanel = new QRCodePanel();
                _qrCodePanel.AddQRCodeRequested += AddQRCodeToSlide;
                _qrCodePane = this.CustomTaskPanes.Add(_qrCodePanel, "二维码生成器");
                _qrCodePane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                _qrCodePane.Width = 300;
            }
            _qrCodePane.Visible = !_qrCodePane.Visible;
        }
        private void AddQRCodeToSlide(Bitmap qrCodeImage)
        {
            try
            {
                var slide = Application.ActiveWindow.View.Slide;

                // 将Bitmap保存到临时文件
                string tempFile = Path.GetTempFileName() + ".png";
                qrCodeImage.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);

                // 添加到幻灯片
                slide.Shapes.AddPicture(tempFile,Microsoft.Office.Core.MsoTriState.msoFalse,Microsoft.Office.Core.MsoTriState.msoTrue,100, 100, qrCodeImage.Width, qrCodeImage.Height);

                // 删除临时文件
                File.Delete(tempFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加二维码失败: {ex.Message}");
            }
        }
    }
}