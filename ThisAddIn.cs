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

//第一作者：yuwenhui2020
//项目地址：https://github.com/yuwenhui2020
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
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        private bool _isInSlideShow = false;
        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;//添加网页的主面板
        private WebViewPanel _webViewPanel;//添加网页的主面板

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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
        
        private System.Drawing.Rectangle GetSlideShowScreenBounds()
        {
            if (Application.SlideShowWindows.Count == 0)
                return Screen.PrimaryScreen.Bounds;

            try
            {
                IntPtr slideShowHwnd = (IntPtr)Application.SlideShowWindows[1].HWND;
                GetWindowRect(slideShowHwnd, out RECT windowRect);

                // 转换为Rectangle
                var windowBounds = new System.Drawing.Rectangle(
                    windowRect.Left, windowRect.Top,
                    windowRect.Right - windowRect.Left,
                    windowRect.Bottom - windowRect.Top);

                // 找到包含大部分窗口的显示器
                foreach (Screen screen in Screen.AllScreens)
                {
                    if (screen.Bounds.IntersectsWith(windowBounds))
                    {
                        return screen.Bounds;
                    }
                }

                return Screen.PrimaryScreen.Bounds;
            }
            catch
            {
                return Screen.PrimaryScreen.Bounds;
            }
        }
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
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
            _slideMonitorTimer.Interval = 200; // 每200毫秒检查一次当前幻灯片页面是否存在网页占位符
            _slideMonitorTimer.Tick += (sender, e) => CheckCurrentSlide();
            _slideMonitorTimer.Start();
        }
        private void CheckCurrentSlide()
        {
            try
            {
                if (!_isInSlideShow || Application.SlideShowWindows.Count == 0)
                {
                    // 不在放映状态，隐藏所有网页
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

                // 如果幻灯片没有变化，不做处理
                if (_currentSlide != null && _currentSlide.SlideID == currentSlide.SlideID)
                    return;

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

                // 如果没有网页占位符，隐藏所有网页
                if (!hasWebView)
                {
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
            foreach (var form in _activeWebViews.Values)
            {
                try
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
                catch (Exception ex)
                {
                    //System.Diagnostics.Debug.WriteLine($"隐藏WebView时出错: {ex.Message}");
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
                //System.Diagnostics.Debug.WriteLine($"OnIdleShowWebView error: {ex.Message}");
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
                shape.Line.Visible = MsoTriState.msoTrue;
                shape.Line.Weight = 1.00f;

                // 只在幻灯片放映时创建WebView
                if (_isInSlideShow && Application.SlideShowWindows.Count > 0)
                {
                    var webForm = new WebBrowserForm();
                    webForm.AddWebView(url, width, height);
                    _activeWebViews[slide] = webForm;
                    ShowWebViewForCurrentSlide();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"添加网页失败: {ex.Message}");
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
        private void ShowWebViewForCurrentSlide()
        {
            if (!_isInSlideShow || Application.SlideShowWindows.Count == 0) return;

            try
            {
                // 隐藏所有已显示的网页窗体
                foreach (var form in _activeWebViews.Values)
                {
                    form.Hide();
                }

                var slideShowWindow = Application.SlideShowWindows[1];
                var currentSlide = slideShowWindow.View.Slide;
                // 设置父窗口
                var hwnd = (IntPtr)slideShowWindow.HWND;
                // 获取放映窗口所在的屏幕边界
                var screenBounds = GetSlideShowScreenBounds();
                // 检查当前幻灯片是否有网页控件
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in currentSlide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_"))
                    {
                        string url = shape.Tags["WebViewUrl"];
                        if (!string.IsNullOrEmpty(url))
                        {
                            int Formwidth = int.Parse(shape.Tags["WebViewWidth"]);
                            int Formheight = int.Parse(shape.Tags["WebViewHeight"]);

                            if (!_activeWebViews.TryGetValue(currentSlide, out WebBrowserForm currentForm))
                            {
                                currentForm = new WebBrowserForm();
                                currentForm.AddWebView(url, Formwidth, Formheight);
                                _activeWebViews[currentSlide] = currentForm;
                            }

                            // 设置为子窗口
                            currentForm.SetAsChildWindow(hwnd);

                            // 获取系统缩放比例并计算缩放比visions的值
                            float visions = GetVisionsBasedOnDpiScaling();

                            // 获取屏幕显示区域尺寸(像素)
                            float screenWidth = slideShowWindow.Width;
                            float screenHeight = slideShowWindow.Height;
                            float screenCenterX = screenWidth / 2;
                            float screenCenterY = screenHeight / 2;

                            // 获取幻灯片页面尺寸(以磅为单位)
                            float slideWidth = currentSlide.Parent.PageSetup.SlideWidth;
                            float slideHeight = currentSlide.Parent.PageSetup.SlideHeight;
                            float slideCenterX = slideWidth / 2;
                            float slideCenterY = slideHeight / 2;

                            // 计算缩放比例(基于最小比例保证完整显示)
                            float scale = Math.Min(screenWidth / slideWidth, screenHeight / slideHeight);

                            // 计算控件中心点相对于幻灯片中心的偏移量(磅)
                            float shapeCenterX = shape.Left + shape.Width / 2;
                            float shapeCenterY = shape.Top + shape.Height / 2;
                            float offsetFromSlideCenterX = shapeCenterX - slideCenterX;
                            float offsetFromSlideCenterY = shapeCenterY - slideCenterY;

                            // 转换为屏幕坐标(从屏幕中心开始计算)
                            float screenShapeCenterX = (screenCenterX + offsetFromSlideCenterX * scale) * visions;
                            float screenShapeCenterY = (screenCenterY + offsetFromSlideCenterY * scale) * visions;

                            // 计算窗体位置(左上角坐标)
                            int left = (int)(screenShapeCenterX - shape.Width * scale * visions / 2);
                            int top = (int)(screenShapeCenterY - shape.Height * scale * visions / 2);
                            int width = (int)(shape.Width * scale * visions);
                            int height = (int)(shape.Height * scale * visions);

                            // 调整坐标到放映屏幕(适配多显示器)
                            left = left + screenBounds.Left;
                            top = top + screenBounds.Top;

                            // 设置窗体位置和大小
                            currentForm.Location = new System.Drawing.Point(left, top);
                            currentForm.Size = new System.Drawing.Size(width, height);

                            // 确保窗体可见
                            currentForm.Show();
                            // 仅设置窗口在父窗口中的Z序
                            SetWindowPos(currentForm.Handle, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                            // 设置焦点到WebView2控件
                            currentForm.webView21.Focus();
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"显示网页失败: {ex.Message}");
            }
        }
        private float GetVisionsBasedOnDpiScaling()
        {
            //根据DPI缩放比例计算visions值
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                // 获取系统DPI与标准96DPI的比例
                float dpiScale = g.DpiX / (48.0f * 1.5f);
                // 因为经过实践发现与屏幕缩放存在2/1.5倍的缩放关系（其实就是我手动输入一次次试出来的）
                return dpiScale;
            }
        }
        private void Application_SlideShowEnd(Presentation Pres)
        {
            _isInSlideShow = false;
            // 停止监控
            if (_slideMonitorTimer != null)
            {
                _slideMonitorTimer.Stop();
                _slideMonitorTimer.Dispose();
                _slideMonitorTimer = null;
            }

            ForceHideAllWebViews();
            _currentSlide = null;
            // 清理所有网页窗体
            foreach (var form in _activeWebViews.Values)
            {
                form.Hide();
                form.Dispose();
            }
            _activeWebViews.Clear();
            CleanupWebView2Data();
        }
        private void ForceHideAllWebViews()
        {
            foreach (var form in _activeWebViews.Values)
            {
                try
                {
                    if (!form.IsDisposed)
                    {
                        if (form.InvokeRequired)
                        {
                            form.Invoke(new Action(() =>
                            {
                                form.Visible = false;
                                form.Close();
                            }));
                        }
                        else
                        {
                            form.Visible = false;
                            form.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"强制隐藏WebView时出错: {ex.Message}");
                }
            }
            _activeWebViews.Clear();
        }
        private void CleanupWebView2Data()
        {
            try
            {
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointAddIn",
                    "WebView2Data");

                if (Directory.Exists(appDataPath))
                {
                    Directory.Delete(appDataPath, true); // 完全删除旧数据
                }
            }
            catch { /* 静默处理清理错误 */ }
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


            // 清理WebView2数据
            CleanupWebView2Data();
        }
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        // 鼠标事件API
        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;


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
    }
}