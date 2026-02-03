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

                // 如果没有网页占位符，隐藏所有网页（包括当前幻灯片的）
                if (!hasWebView)
                {
                    if (_activeWebViews.ContainsKey(currentSlide))
                    {
                        var formToRemove = _activeWebViews[currentSlide];
                        formToRemove.Hide();
                        formToRemove.Dispose();
                        _activeWebViews.Remove(currentSlide);
                    }
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
            List<Slide> slidesToRemove = new List<Slide>();

            foreach (var kvp in _activeWebViews)
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
                _activeWebViews.Remove(slide);
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

                // 获取放映窗口句柄
                var hwnd = (IntPtr)slideShowWindow.HWND;

                // 获取窗口矩形
                GetWindowRect(hwnd, out RECT windowRect);
                float screenWidth = windowRect.Right - windowRect.Left;
                float screenHeight = windowRect.Bottom - windowRect.Top;

                // 先隐藏所有网页（不仅仅是当前幻灯片的）
                HideAllWebViewsExcept(currentSlide);

                // 检查当前幻灯片是否有网页控件
                bool hasWebView = false;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in currentSlide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_"))
                    {
                        string url = shape.Tags["WebViewUrl"];
                        if (!string.IsNullOrEmpty(url))
                        {
                            hasWebView = true;
                            int Formwidth = int.Parse(shape.Tags["WebViewWidth"]);
                            int Formheight = int.Parse(shape.Tags["WebViewHeight"]);

                            // 获取或创建WebView窗体
                            if (!_activeWebViews.TryGetValue(currentSlide, out WebBrowserForm currentForm))
                            {
                                currentForm = new WebBrowserForm();
                                currentForm.FormClosed += (s, e) =>
                                {
                                    if (_activeWebViews.ContainsKey(currentSlide))
                                        _activeWebViews.Remove(currentSlide);
                                };

                                // 初始化WebView
                                currentForm.AddWebView(url, Formwidth, Formheight);
                                _activeWebViews[currentSlide] = currentForm;
                            }
                            else
                            {
                                // 如果窗体已存在但可能已关闭，重新创建
                                if (currentForm.IsDisposed)
                                {
                                    currentForm = new WebBrowserForm();
                                    currentForm.FormClosed += (s, e) =>
                                    {
                                        if (_activeWebViews.ContainsKey(currentSlide))
                                            _activeWebViews.Remove(currentSlide);
                                    };
                                    currentForm.AddWebView(url, Formwidth, Formheight);
                                    _activeWebViews[currentSlide] = currentForm;
                                }
                                else
                                {
                                    if (currentForm.CurrentUrl != url)
                                    {
                                        currentForm.NavigateToUrl(url);
                                    }
                                }
                            }

                            // 设置窗体为子窗口
                            if (currentForm.InvokeRequired)
                            {
                                currentForm.Invoke(new Action(() =>
                                {
                                    currentForm.SetAsChildWindow(hwnd);
                                }));
                            }
                            else
                            {
                                currentForm.SetAsChildWindow(hwnd);
                            }

                            // 计算位置和大小
                            CalculateAndSetWebViewPosition(currentForm, shape, currentSlide, screenBounds, screen, hwnd);

                            break; // 只处理第一个找到的WebView
                        }
                    }
                }

                // 如果没有找到WebView，确保清理当前幻灯片的网页
                if (!hasWebView && _activeWebViews.ContainsKey(currentSlide))
                {
                    var formToRemove = _activeWebViews[currentSlide];
                    formToRemove.Hide();
                    formToRemove.Dispose();
                    _activeWebViews.Remove(currentSlide);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示网页失败: {ex.Message}");
            }
        }
        private void HideAllWebViewsExcept(Slide exceptSlide)
        {
            List<Slide> slidesToRemove = new List<Slide>();

            foreach (var kvp in _activeWebViews)
            {
                if (kvp.Key.SlideID != exceptSlide.SlideID)
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
                                    form.Hide();
                                }));
                            }
                            else
                            {
                                form.Hide();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"隐藏网页失败: {ex.Message}");
                        slidesToRemove.Add(kvp.Key);
                    }
                }
            }

            // 从字典中移除已处理的幻灯片（仅当出现异常时）
            foreach (var slide in slidesToRemove)
            {
                _activeWebViews.Remove(slide);
            }
        }

        private void CalculateAndSetWebViewPosition(WebBrowserForm form, Microsoft.Office.Interop.PowerPoint.Shape shape,
            Slide currentSlide, System.Drawing.Rectangle screenBounds, Screen screen, IntPtr hwnd)
        {
            try
            {
                // 获取显示器DPI缩放
                float dpiScale = GetDpiScaleForScreen(screen);

                // 获取幻灯片页面尺寸(磅)
                float slideWidth = currentSlide.Parent.PageSetup.SlideWidth;
                float slideHeight = currentSlide.Parent.PageSetup.SlideHeight;

                // 获取放映窗口的实际像素尺寸
                GetWindowRect(hwnd, out RECT windowRect);
                float windowWidth = windowRect.Right - windowRect.Left;
                float windowHeight = windowRect.Bottom - windowRect.Top;

                // 计算放映窗口的缩放比例
                float scaleX = windowWidth / slideWidth;
                float scaleY = windowHeight / slideHeight;
                float scale = Math.Min(scaleX, scaleY); // 使用最小比例确保完整显示

                // 计算形状在放映窗口中的位置
                // 首先计算幻灯片在窗口中的偏移（居中显示）
                float slideOffsetX = (windowWidth - slideWidth * scale) / 2;
                float slideOffsetY = (windowHeight - slideHeight * scale) / 2;

                // 计算形状位置（相对于窗口）
                int left = (int)(slideOffsetX + shape.Left * scale);
                int top = (int)(slideOffsetY + shape.Top * scale);
                int width = (int)(shape.Width * scale);
                int height = (int)(shape.Height * scale);

                // 应用DPI缩放
                left = (int)(left * dpiScale);
                top = (int)(top * dpiScale);
                width = (int)(width * dpiScale);
                height = (int)(height * dpiScale);

                // 调整到屏幕坐标
                left += screenBounds.Left;
                top += screenBounds.Top;

                // 确保尺寸合理
                if (width < 10) width = 10;
                if (height < 10) height = 10;

                // 设置窗体位置和大小
                if (form.InvokeRequired)
                {
                    form.Invoke(new Action(() =>
                    {
                        form.Location = new System.Drawing.Point(left, top);
                        form.Size = new System.Drawing.Size(width, height);
                        form.Show();
                        SetWindowPos(form.Handle, HWND_TOP, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    }));
                }
                else
                {
                    form.Location = new System.Drawing.Point(left, top);
                    form.Size = new System.Drawing.Size(width, height);
                    form.Show();
                    SetWindowPos(form.Handle, HWND_TOP, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"计算位置失败: {ex.Message}");
            }
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

            // 停止监控
            if (_slideMonitorTimer != null)
            {
                _slideMonitorTimer.Stop();
                _slideMonitorTimer.Dispose();
                _slideMonitorTimer = null;
            }

            ForceHideAllWebViews();
            _currentSlide = null;
        }
        private void ForceHideAllWebViews()
        {
            List<Slide> slidesToRemove = new List<Slide>();

            foreach (var kvp in _activeWebViews)
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
                _activeWebViews.Remove(slide);
            }
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

            // 停止所有Web控制台活动
            WebControl.StopAll();

            // 清理WebView2数据
            CleanupWebView2Data();
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