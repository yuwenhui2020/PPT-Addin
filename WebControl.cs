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
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointAddIn
{
    public static class WebControl
    {
        private static HttpListener _httpListener;
        private static Thread _serverThread;
        private static WebSocket _screenWebSocket;
        private static bool _isScreenSharing = false;
        private static Thread _screenCaptureThread;
        private static readonly object _screenLock = new object();
        private static Bitmap _lastScreenCapture;
        private static int _screenCaptureInterval = 50; // 毫秒
        private static int _serverPort = 8888;//web控制台的端口
        private static int _screenSharePort = 8887;//屏幕共享
        private static int _textInputPort = 8889;//控制输入
        private static bool _webConsoleEnabled = false;

        public static void ToggleWebConsole()
        {
            if (_webConsoleEnabled)
            {
                StopWebServer();
                _webConsoleEnabled = false;
            }
            else
            {
                StartWebServer();
                _webConsoleEnabled = true;
            }
        }

        private static void StartWebServer()
        {
            try
            {
                if (_httpListener != null && _httpListener.IsListening)
                    return;


                // 配置URL ACL（需要管理员权限）
                ConfigureUrlAcl();

                // 启动HTTP监听器
                _httpListener = new HttpListener();
                _httpListener.Prefixes.Add($"http://+:{_serverPort}/");
                _httpListener.Prefixes.Add($"http://+:{_screenSharePort}/");
                _httpListener.Prefixes.Add($"http://+:{_textInputPort}/");
                
                
                _httpListener.Start();
                _serverThread = new Thread(ListenForRequests)
                {
                    IsBackground = true,
                    Priority = ThreadPriority.Normal
                };
                _serverThread.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动Web服务器失败: {ex.Message}\n\n可能需要以管理员身份运行PowerPoint\n或者端口 {_serverPort}/{_screenSharePort}/{_textInputPort} 已被占用，请关闭占用程序后重试",
                    "服务器启动错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                StopWebServer();
            }
        }
        private static bool IsPortAvailable(int port)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    client.Connect("127.0.0.1", port);
                    return false;
                }
            }
            catch
            {
                return true;
            }
        }
        public static void StopAll()
        {
            if (_webConsoleEnabled)
            {
                StopWebServer();
                _webConsoleEnabled = false;
            }

            // 停止屏幕共享
            _isScreenSharing = false;
            _screenWebSocket?.Dispose();
            _screenWebSocket = null;

            if (_screenCaptureThread != null && _screenCaptureThread.IsAlive)
            {
                _screenCaptureThread.Join(1000); // 等待最多1秒
            }

            lock (_screenLock)
            {
                _lastScreenCapture?.Dispose();
                _lastScreenCapture = null;
            }
        }

        private static void ConfigureUrlAcl()
        {
            try
            {
                // 删除旧规则（如果存在）
                ExecuteCommand($"netsh http delete urlacl url=http://+:{_serverPort}/", true);
                ExecuteCommand($"netsh http delete urlacl url=http://+:{_screenSharePort}/", true);
                ExecuteCommand($"netsh http delete urlacl url=http://+:{_textInputPort}/", true);
                
                // 添加新规则
                ExecuteCommand($"netsh http add urlacl url=http://+:{_serverPort}/ user={Environment.UserName}", true);
                ExecuteCommand($"netsh http add urlacl url=http://+:{_screenSharePort}/ user={Environment.UserName}", true);
                ExecuteCommand($"netsh http add urlacl url=http://+:{_textInputPort}/ user={Environment.UserName}", true);
                
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"配置URL ACL时出错: {ex.Message}");
                throw new Exception("需要管理员权限配置网络访问。请右键点击PowerPoint，选择'以管理员身份运行'");
            }
        }

        private static string ExecuteCommand(string command, bool requireAdmin = false)
        {
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c " + command,
                    CreateNoWindow = true,
                    UseShellExecute = requireAdmin,
                    Verb = requireAdmin ? "runas" : "", // 需要管理员权限时触发UAC
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardError = !requireAdmin,
                    RedirectStandardOutput = !requireAdmin
                };

                using (var process = Process.Start(processInfo))
                {
                    if (!requireAdmin)
                    {
                        string output = process.StandardOutput.ReadToEnd();
                        string error = process.StandardError.ReadToEnd();
                        process.WaitForExit();

                        if (process.ExitCode != 0)
                        {
                            throw new Exception($"命令执行失败({process.ExitCode}): {error}");
                        }

                        return output;
                    }
                    else
                    {
                        process.WaitForExit();
                        return process.ExitCode == 0 ? "成功" : "失败";
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"执行命令失败: {command} - {ex.Message}");
                throw;
            }
        }

        private static void StopWebServer()
        {
            _httpListener?.Stop();
            _serverThread?.Join();

            // 使用FirewallHelper移除防火墙规则
            //每次关闭都移除防火墙太麻烦了，先不删除用着看看有没有问题
            /*
            FirewallHelper.RemoveFirewallRule(_serverPort);
            FirewallHelper.RemoveFirewallRule(_screenSharePort);
            FirewallHelper.RemoveFirewallRule(_textInputPort);
            */
        }

        private static void ListenForRequests()
        {
            while (_httpListener.IsListening)
            {
                try
                {
                    var context = _httpListener.GetContext();

                    if (context.Request.IsWebSocketRequest)
                    {
                        // 处理WebSocket请求
                        Task.Run(async () => await HandleWebSocketRequest(context)).ContinueWith(t =>
                        {
                            if (t.Exception != null)
                            {
                                Debug.WriteLine($"WebSocket处理错误: {t.Exception}");
                            }
                        });
                    }
                    else
                    {
                        // 处理普通HTTP请求
                        ProcessRequest(context);
                    }
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"{ex.Message}");
                }
            }
        }

        private static void ProcessRequest(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                if (request.HttpMethod == "GET" && request.Url.AbsolutePath == "/control")
                {
                    var command = request.QueryString["cmd"];
                    if (command == "status")
                    {
                        HandleStatusCommand(context);
                        return;
                    }
                }
                else if (request.HttpMethod == "POST" && request.Url.AbsolutePath == "/control")
                {
                    HandlePostCommand(context);
                    return;
                }

                // 返回控制台HTML页面
                var html = GetControlPageHtml();
                var buffer = System.Text.Encoding.UTF8.GetBytes(html);
                response.ContentLength64 = buffer.Length;
                response.ContentType = "text/html";
                response.OutputStream.Write(buffer, 0, buffer.Length);
            }
            finally
            {
                response.Close();
            }
        }

        private static void HandlePostCommand(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                {
                    var json = reader.ReadToEnd();
                    var data = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                    if (data != null && data.TryGetValue("cmd", out var command))
                    {
                        string result = "";
                        switch (command.ToLower())
                        {
                            case "start":
                                result = StartPresentation();
                                break;
                            case "stop":
                                result = StopPresentation();
                                break;
                            case "next":
                            case "prev":
                            case "first":
                            case "last":
                                result = HandleSlideNavigation(command);
                                break;
                            default:
                                throw new ArgumentException("未知命令");
                        }

                        var buffer = System.Text.Encoding.UTF8.GetBytes(result);
                        response.ContentLength64 = buffer.Length;
                        response.ContentType = "application/json";
                        response.OutputStream.Write(buffer, 0, buffer.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                response.StatusCode = 400;
                var buffer = System.Text.Encoding.UTF8.GetBytes(ex.Message);
                response.ContentLength64 = buffer.Length;
                response.OutputStream.Write(buffer, 0, buffer.Length);
            }
        }

        private static string HandleSlideNavigation(string command)
        {
            if (Globals.ThisAddIn.Application.SlideShowWindows.Count == 0)
            {
                throw new InvalidOperationException("未在放映");
            }

            var view = Globals.ThisAddIn.Application.SlideShowWindows[1].View;

            try
            {
                switch (command.ToLower())
                {
                    case "next": view.Next(); break;
                    case "prev": view.Previous(); break;
                    case "first": view.First(); break;
                    case "last": view.Last(); break;
                }
                return "OK";
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"{ex.Message}");
                throw new InvalidOperationException("未在放映");
            }
        }
        
        private static string StartPresentation()
        {
            try
            {
                if (Globals.ThisAddIn.Application.SlideShowWindows.Count > 0)
                {
                    return "已在放映";
                }

                if (Globals.ThisAddIn.Application.ActivePresentation == null)
                {
                    throw new InvalidOperationException("PowerPoint未打开演示文稿");
                }

                Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.Run();
                return "OK";
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"开始放映失败: {ex.Message}");
            }
        }

        private static string StopPresentation()
        {
            if (Globals.ThisAddIn.Application.SlideShowWindows.Count == 0)
            {
                return "未在放映";
            }
            Globals.ThisAddIn.Application.SlideShowWindows[1].View.Exit();
            return "OK";
        }

        private static void HandleStatusCommand(HttpListenerContext context)
        {
            var response = context.Response;
            var status = new
            {
                isInSlideshow = Globals.ThisAddIn.Application.SlideShowWindows.Count > 0,
                currentSlide = Globals.ThisAddIn.Application.SlideShowWindows.Count > 0 ?
                    Globals.ThisAddIn.Application.SlideShowWindows[1].View.Slide.SlideIndex : 0,
                totalSlides = Globals.ThisAddIn.Application.SlideShowWindows.Count > 0 ?
                    Globals.ThisAddIn.Application.SlideShowWindows[1].View.Slide.Parent.Slides.Count : 0
            };

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(status);
            var buffer = System.Text.Encoding.UTF8.GetBytes(json);

            response.ContentLength64 = buffer.Length;
            response.ContentType = "application/json";
            response.OutputStream.Write(buffer, 0, buffer.Length);
        }

        private static string GetControlPageHtml()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "PowerPointAddIn.WebConsole.index.html";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    MessageBox.Show("无法加载Web控制台界面");
                    return "<h1>Error: Web控制台界面未找到</h1>";
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public static string GetConsoleUrl()
        {
            return $"http://{GetLocalIPAddress()}:{_serverPort}";
        }

        public static bool IsConsoleEnabled()
        {
            return _webConsoleEnabled;
        }

        private static string GetLocalIPAddress()
        {
            try
            {
                // 获取所有IPv4地址，排除回环和虚拟适配器
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork &&
                        !ip.ToString().StartsWith("169.254.")) // 排除APIPA地址
                    {
                        return ip.ToString();
                    }
                }

                // 如果没有找到合适的IP，返回第一个IPv4地址
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        return ip.ToString();
                    }
                }
            }
            catch
            {
                // 如果获取失败
            }
            return "localhost";
        }

        private static async Task SendScreenInfo(WebSocket webSocket)
        {
            try
            {
                int totalWidth = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
                int totalHeight = System.Windows.Forms.SystemInformation.VirtualScreen.Height;
                int screenLeft = System.Windows.Forms.SystemInformation.VirtualScreen.Left;
                int screenTop = System.Windows.Forms.SystemInformation.VirtualScreen.Top;

                var screenInfo = new
                {
                    action = "screenInfo",
                    width = totalWidth,
                    height = totalHeight,
                    offsetX = screenLeft,
                    offsetY = screenTop,
                    captureWidth = 1280, // 与CaptureScreen中的scaledWidth一致
                    captureHeight = (int)(totalHeight * ((double)1280 / totalWidth))
                };

                var json = Newtonsoft.Json.JsonConvert.SerializeObject(screenInfo);
                var buffer = System.Text.Encoding.UTF8.GetBytes(json);

                await webSocket.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Text, true, CancellationToken.None);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送屏幕信息错误: {ex.Message}");
            }
        }

        private static async Task HandleWebSocketRequest(HttpListenerContext context)
        {
            try
            {
                var webSocketContext = await context.AcceptWebSocketAsync(null);
                var webSocket = webSocketContext.WebSocket;

                // 根据端口区分功能
                if (context.Request.LocalEndPoint.Port == _screenSharePort)
                {
                    _screenWebSocket = webSocket;
                    _isScreenSharing = true;
                    await SendScreenInfo(_screenWebSocket);

                    if (_screenCaptureThread == null || !_screenCaptureThread.IsAlive)
                    {
                        _screenCaptureThread = new Thread(ScreenCaptureLoop);
                        _screenCaptureThread.IsBackground = true;
                        _screenCaptureThread.Start();
                    }

                    await HandleWebSocketMessages(webSocket);
                }
                else if (context.Request.LocalEndPoint.Port == _textInputPort)
                {
                    await HandleTextInputWebSocket(webSocket);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebSocket连接错误: {ex.Message}");
            }
        }
        private static async Task HandleTextInputWebSocket(WebSocket webSocket)
        {
            try
            {
                var buffer = new byte[1024];
                while (webSocket.State == WebSocketState.Open)
                {
                    var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);

                    if (result.MessageType == WebSocketMessageType.Text)
                    {
                        var message = System.Text.Encoding.UTF8.GetString(buffer, 0, result.Count);
                        var data = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(message);

                        if (data.action == "textInput")
                        {
                            string text = data.text;
                            await SendTextToActiveWindow(text);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"文本输入WebSocket处理错误: {ex.Message}");
            }
        }

        private static async Task HandleWebSocketMessages(WebSocket webSocket)
        {
            var buffer = new byte[1024];
            try
            {
                while (webSocket != null && webSocket.State == WebSocketState.Open)
                {
                    var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);

                    if (result.MessageType == WebSocketMessageType.Text)
                    {
                        var message = System.Text.Encoding.UTF8.GetString(buffer, 0, result.Count);
                        try
                        {
                            var data = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(message);

                            // 处理点击事件
                            if (data.action == "click")
                            {
                                int x = data.x;
                                int y = data.y;

                                // 在主线程执行点击操作
                                System.Windows.Forms.Application.DoEvents();
                                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(x, y);
                                MouseClick();
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"WebSocket消息处理错误: {ex.Message}");
                        }
                    }
                    else if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WebSocket通信错误: {ex.Message}");
            }
            finally
            {
                if (webSocket == _screenWebSocket)
                {
                    _isScreenSharing = false;
                    _screenWebSocket?.Dispose();
                    _screenWebSocket = null;
                }
                else
                {
                    webSocket?.Dispose();
                }
            }
        }

        //文本输入功能
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
            public static int Size => Marshal.SizeOf(typeof(INPUT));
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        private const uint WM_CHAR = 0x0102;
        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYDOWN = 0x0000;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        private static async Task SendTextToActiveWindow(string text)
        {
            try
            {
                // 获取当前活动窗口句柄
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero)
                {
                    Debug.WriteLine("无法获取活动窗口");
                    return;
                }

                // 使用SendInput模拟键盘输入
                foreach (char c in text)
                {
                    // 发送WM_CHAR消息
                    SendMessage(hwnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);

                    // 同时使用SendInput模拟实际按键
                    var inputs = new INPUT[2];

                    // 按下
                    inputs[0].type = INPUT_KEYBOARD;
                    inputs[0].U.ki.wScan = (ushort)c;
                    inputs[0].U.ki.dwFlags = KEYEVENTF_UNICODE;

                    // 释放
                    inputs[1].type = INPUT_KEYBOARD;
                    inputs[1].U.ki.wScan = (ushort)c;
                    inputs[1].U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

                    SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));

                    await Task.Delay(10); // 小延迟防止输入过快
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送文本失败: {ex.Message}");
            }
        }

        //文本输入功能结束

        private static void ScreenCaptureLoop()
        {
            try
            {
                while (_isScreenSharing)
                {
                    try
                    {
                        var screen = CaptureScreen();

                        lock (_screenLock)
                        {
                            _lastScreenCapture?.Dispose();
                            _lastScreenCapture = screen;
                        }

                        if (_screenWebSocket != null && _screenWebSocket.State == WebSocketState.Open)
                        {
                            using (var ms = new MemoryStream())
                            {
                                screen.Save(ms, ImageFormat.Jpeg);
                                var buffer = ms.ToArray();
                                _screenWebSocket.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Binary, true, CancellationToken.None).Wait();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"屏幕捕获错误: {ex.Message}");
                    }

                    Thread.Sleep(_screenCaptureInterval);
                }
            }
            finally
            {
                lock (_screenLock)
                {
                    _lastScreenCapture?.Dispose();
                    _lastScreenCapture = null;
                }
            }
        }

        private static Bitmap CaptureScreen()
        {
            // 获取所有屏幕的总体边界
            int totalWidth = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
            int totalHeight = System.Windows.Forms.SystemInformation.VirtualScreen.Height;
            int screenLeft = System.Windows.Forms.SystemInformation.VirtualScreen.Left;
            int screenTop = System.Windows.Forms.SystemInformation.VirtualScreen.Top;

            var bitmap = new Bitmap(totalWidth, totalHeight);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(screenLeft, screenTop, 0, 0, new System.Drawing.Size(totalWidth, totalHeight));
            }

            // 调整大小以减小带宽使用
            var scaledWidth = 1280;
            var scaledHeight = (int)(totalHeight * ((double)scaledWidth / totalWidth));

            var scaledBitmap = new Bitmap(scaledWidth, scaledHeight);
            using (var g = Graphics.FromImage(scaledBitmap))
            {
                g.DrawImage(bitmap, 0, 0, scaledWidth, scaledHeight);
            }

            bitmap.Dispose();
            return scaledBitmap;
        }

        private static void MouseClick()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }

        // 鼠标事件API
        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;


    }
}