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
using System.Net.NetworkInformation;
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
        // =====================================================================
        // 端口常量：全项目唯一定义处。
        // 原实现在 WebControl.cs 和 WebViewPanel.cs 各定义一份，改一处忘另一处
        // 就会出现“防火墙放行的端口”和“实际监听的端口”不一致。现统一在此暴露，
        // WebViewPanel 改为引用本类的静态只读数组。
        // =====================================================================
        public const int ServerPort = 8888;      // WEB 控制台页面（首选）
        public const int ScreenSharePort = 8887; // 屏幕共享 (WebSocket)（首选）
        public const int TextInputPort = 8889;   // 文本/控制输入 (WebSocket)（首选）

        // 当前实际使用的端口。默认取上面的首选值；若首选被占用（如第三方软件占用 8888），
        // StartWebServer 会自动整体后移并改写这里。WebSocket 路由依据 LocalEndPoint.Port
        // 区分功能，因此必须引用这里的实时值，不能再用 const。
        private static int _serverPort = ServerPort;
        private static int _screenSharePort = ScreenSharePort;
        private static int _textInputPort = TextInputPort;

        // 当前实际端口集合（供面板/防火墙使用）
        public static int[] AllPorts => new[] { _serverPort, _screenSharePort, _textInputPort };

        /// <summary>
        /// 本次会话的访问码。
        ///
        /// 安全修复：此前 Web 控制台对所有请求完全不鉴权，且 HttpListener 绑定的是
        /// “所有网卡(http://+/)”，配合安装脚本自动放行的防火墙规则，意味着任何能访问
        /// 到本机 8888/8887/8889 端口的人都可以：
        ///   * 通过 /control 远程翻页、启停放映；
        ///   * 通过 8889 端口用 SendInput 向当前前台窗口注入任意按键（等同远程键盘）；
        ///   * 通过 8887 端口持续拉取整屏截图（等同远程桌面）。
        /// 现引入随机访问码：每次开启控制台生成新的 6 位数字，必须随请求提供，
        /// 否则一律 401 拒绝 / 拒绝 WebSocket 握手。
        /// </summary>
        public static string AccessCode { get; private set; }
        private static readonly object _codeLock = new object();

        private static bool _webConsoleEnabled = false;

        /// <summary>
        /// 生成 6 位随机数字访问码。
        /// </summary>
        private static string GenerateAccessCode()
        {
            int value;
            lock (_codeLock)
            {
                using (var rng = new System.Security.Cryptography.RNGCryptoServiceProvider())
                {
                    var bytes = new byte[4];
                    rng.GetBytes(bytes);
                    value = Math.Abs(BitConverter.ToInt32(bytes, 0));
                }
            }
            return (value % 1000000).ToString("D6");
        }

        /// <summary>
        /// 校验请求中携带的访问码是否正确。缺失或不匹配均视为未授权。
        /// </summary>
        private static bool IsAuthorized(HttpListenerRequest request)
        {
            string expected = AccessCode;
            if (string.IsNullOrEmpty(expected))
                return false;

            string provided = request.QueryString["code"];
            if (string.IsNullOrEmpty(provided))
                provided = request.Headers["X-Access-Code"];

            return !string.IsNullOrEmpty(provided) &&
                   string.Equals(provided, expected, StringComparison.Ordinal);
        }

        public static void ToggleWebConsole()
        {
            if (_webConsoleEnabled)
            {
                StopWebServer();
                _webConsoleEnabled = false;
            }
            else
            {
                if (StartWebServer())
                {
                    _webConsoleEnabled = true;
                }
            }
        }

        /// <summary>
        /// 启动 Web 控制台服务。
        /// 返回 true 表示启动成功；false 表示失败（已提示用户）。
        /// </summary>
        private static bool StartWebServer()
        {
            try
            {
                if (_httpListener != null && _httpListener.IsListening)
                    return true;

                // 1) 选定可用端口：首选 8887-8889，若被占用（如第三方软件占用 8888）则整体后移，
                //    确保控制台一定能启动，不会再因“端口被占用”而失败。
                if (!TryAllocatePorts(out int srv, out int scr, out int txt))
                {
                    MessageBox.Show(
                        "无法找到可用的端口组合（已尝试 8887-8889 及后续多组均被占用）。\n\n" +
                        "请关闭占用端口的其它程序（如某些网速/网关类软件）后重试。",
                        "服务器启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                _serverPort = srv; _screenSharePort = scr; _textInputPort = txt;

                // 2) 配置 URL 保留项(urlacl)与防火墙：按需提权，最多 1 次 UAC，已配置过时 0 次。
                if (!EnsureNetworkAccess())
                {
                    return false;
                }

                if (!IsHttpListenerSupported())
                {
                    MessageBox.Show("当前 Windows 版本不支持 HttpListener，无法启用 Web 控制台。",
                        "服务器启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 每次启动重新生成随机访问码，防止上一次会话的访问码被复用
                AccessCode = GenerateAccessCode();

                _httpListener = new HttpListener();
                _httpListener.Prefixes.Add($"http://+:{_serverPort}/");
                _httpListener.Prefixes.Add($"http://+:{_screenSharePort}/");
                _httpListener.Prefixes.Add($"http://+:{_textInputPort}/");

                try
                {
                    _httpListener.Start();
                }
                catch (HttpListenerException hlEx)
                {
                    // 最常见原因：urlacl 没配成功（未管理员初始化）或端口被占用
                    MessageBox.Show(
                        $"无法监听端口 {_serverPort}/{_screenSharePort}/{_textInputPort}。\n\n" +
                        $"错误：{hlEx.Message}（错误码 {hlEx.ErrorCode}）\n\n" +
                        "常见原因：\n" +
                        "  1) 未以管理员身份做过端口初始化 —— 请右键运行 InitWebPorts.cmd；\n" +
                        "  2) 端口已被其它程序占用。",
                        "服务器启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    StopWebServer();
                    return false;
                }

                _serverThread = new Thread(ListenForRequests)
                {
                    IsBackground = true,
                    Priority = ThreadPriority.Normal
                };
                _serverThread.Start();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动Web服务器失败: {ex.Message}",
                    "服务器启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                StopWebServer();
                return false;
            }
        }

        /// <summary>
        /// 老版本 Windows（如 XP/2003 早期版本）不带 HttpListener，提前给出可读提示。
        /// </summary>
        private static bool IsHttpListenerSupported()
        {
            return HttpListener.IsSupported;
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

        /// <summary>
        /// 在首选端口（8887/8888/8889）连续三段中挑选一组三个端口都空闲的组合。
        /// 若首选被占用（例如第三方软件占用 8888），整体后移 3 继续尝试，
        /// 直到找到一组三个连续端口都可用为止。
        /// 约定：screen=base, server=base+1, text=base+2（与历史 8887/8888/8889 一致，
        /// 客户端依此从服务器端口推导 screen/text 端口）。
        /// </summary>
        private static bool TryAllocatePorts(out int server, out int screen, out int text)
        {
            server = screen = text = 0;
            int preferredBase = Math.Min(ScreenSharePort, Math.Min(ServerPort, TextInputPort)); // 8887
            for (int attempt = 0; attempt < 100; attempt++)
            {
                int b = preferredBase + attempt * 3;
                int sc = b;       // screen
                int s = b + 1;    // server
                int t = b + 2;    // text
                if (IsPortFree(s) && IsPortFree(sc) && IsPortFree(t))
                {
                    server = s; screen = sc; text = t;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 判断某个 TCP 端口当前是否空闲（本地无进程监听）。使用 TcpListener 试绑定，
        /// 比解析 netstat 更可靠，且无需管理员权限。
        /// </summary>
        private static bool IsPortFree(int port)
        {
            try
            {
                var listener = new TcpListener(IPAddress.Any, port);
                listener.Start();
                listener.Stop();
                return true;
            }
            catch (SocketException)
            {
                return false;
            }
        }

        /// <summary>
        /// 确保本机已配置好 Web 控制台所需的 URL 保留项(urlacl)与防火墙入站规则。
        ///
        /// 修复点：
        ///  1) 原先对 3 个端口各自启动一个提权 netsh 进程，会弹出 3 次 UAC。
        ///     现改为“一次非提权预检 + 最多一次提权批处理”，正常情况（已配置过）0 次 UAC，
        ///     首次配置仅 1 次 UAC。
        ///  2) 原先固定 user={Environment.UserName}，域环境/含非 ASCII 用户名时 netsh 报错，
        ///     导致配置失败。现统一使用 user=Everyone（控制台另由随机访问码鉴权，安全性不受影响）。
        ///  3) 顺带补齐防火墙入站规则，保证手机等局域网设备能连入（原来只有 urlacl，没开防火墙）。
        /// </summary>
        public static bool EnsureNetworkAccess()
        {
            try
            {
                int[] ports = AllPorts; // 当前实际端口

                bool needElevated = false;
                foreach (int port in ports)
                {
                    if (!IsUrlAclReserved(port) || !IsFirewallRuleExists(port))
                    {
                        needElevated = true;
                        break;
                    }
                }

                if (!needElevated)
                    return true;

                // 生成一次性批处理，urlacl + 防火墙 全部用同一条提权进程完成（仅 1 次 UAC）。
                string batPath = Path.Combine(Path.GetTempPath(),
                    "PPTWebConsole_Init_" + Guid.NewGuid().ToString("N") + ".cmd");
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("@echo off");
                foreach (int port in ports)
                {
                    // urlacl：已存在时 add 会报 183，忽略即可
                    sb.AppendLine($"netsh http add urlacl url=http://+:{port}/ user=Everyone >nul 2>&1");
                    // 防火墙：先删后加，保证幂等（与 FirewallHelper 同名规则一致）
                    sb.AppendLine($"netsh advfirewall firewall delete rule name=\"RemoteControl_{port}\" >nul 2>&1");
                    sb.AppendLine($"netsh advfirewall firewall add rule name=\"RemoteControl_{port}\" dir=in action=allow protocol=TCP localport={port} >nul 2>&1");
                }
                sb.AppendLine("exit /b 0");
                File.WriteAllText(batPath, sb.ToString(), System.Text.Encoding.ASCII);

                var psi = new ProcessStartInfo
                {
                    FileName = batPath,
                    CreateNoWindow = true,
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var p = Process.Start(psi))
                {
                    p.WaitForExit();
                    int code = p.ExitCode;
                    try { File.Delete(batPath); } catch { }
                    if (code != 0)
                        throw new Exception($"网络初始化脚本返回 {code}");
                }

                // 提权脚本执行后再次确认 urlacl 确实生效，给出可读错误
                foreach (int port in ports)
                {
                    if (!IsUrlAclReserved(port))
                        throw new Exception($"端口 {port} 的 URL 保留项配置未生效（可能被安全软件拦截）。");
                }
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"配置网络访问出错: {ex.Message}");
                MessageBox.Show(
                    "需要管理员权限来配置网络访问（URL 保留项 / 防火墙）。\n\n" +
                    $"错误：{ex.Message}\n\n" +
                    "请以管理员身份运行 PowerPoint，或手动右键运行 InitWebPorts.cmd。",
                    "服务器启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 非提权检查某端口是否已有 http urlacl 保留项（读取 netsh 输出即可，无需管理员）。
        /// </summary>
        private static bool IsUrlAclReserved(int port)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "netsh.exe",
                    Arguments = "http show urlacl",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    string outp = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                    return outp.IndexOf($"http://+:{port}/", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }
            catch
            {
                // 读取失败则保守认为“未配置”，交由后续提权脚本处理
                return false;
            }
        }

        /// <summary>
        /// 非提权检查防火墙入站规则是否存在（与 FirewallHelper 同名规则一致）。
        /// </summary>
        private static bool IsFirewallRuleExists(int port)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "netsh.exe",
                    Arguments = $"advfirewall firewall show rule name=\"RemoteControl_{port}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    string outp = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                    return outp.IndexOf("RemoteControl_" + port, StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void StopWebServer()
        {
            try
            {
                if (_httpListener != null)
                {
                    try { _httpListener.Stop(); } catch { }
                    // Close() 释放 HttpListener 占用的句柄；
                    // 少了这一步，反复开关控制台会导致端口对象泄漏、下次 Start 失败。
                    try { _httpListener.Close(); } catch { }
                }
            }
            catch { }
            finally
            {
                _httpListener = null;
            }

            // 等待监听线程退出，但必须带超时：
            // 原实现使用无参 Join()，一旦线程因异常卡住会连带卡死 PowerPoint UI 线程。
            if (_serverThread != null)
            {
                try
                {
                    if (!_serverThread.Join(3000))
                    {
                        Debug.WriteLine("Web控制台监听线程未能在3秒内退出，已放弃等待");
                    }
                }
                catch { }
                _serverThread = null;
            }

            // 防火墙规则不在关闭时移除（与列表其它注释一致：每次删除体验很差）
            AccessCode = null;
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

        /// <summary>
        /// 返回 401 未授权响应。
        /// </summary>
        private static void SendUnauthorized(HttpListenerContext context)
        {
            var response = context.Response;
            try
            {
                response.StatusCode = 401;
                response.ContentType = "text/plain; charset=utf-8";
                var msg = System.Text.Encoding.UTF8.GetBytes("401 Unauthorized: 需要正确的访问码");
                response.ContentLength64 = msg.Length;
                response.OutputStream.Write(msg, 0, msg.Length);
            }
            catch { }
        }

        private static void ProcessRequest(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                // 未授权请求一律拒绝：不带正确访问码的访问直接返回 401，
                // 避免翻页/放映控制、以及后续的 WebSocket 通道被他人使用。
                if (!IsAuthorized(request))
                {
                    SendUnauthorized(context);
                    return;
                }

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
                response.ContentType = "text/html; charset=utf-8";
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

            // 修复：原实现没有异常保护。换页瞬间访问 View.Slide 会抛 COM 异常，
            // 异常冒出去后客户端收不到任何响应，控制台的状态轮询就此卡死不再更新。
            // 现改为：任何异常都返回一个可解析的 JSON，保证轮询链路不中断。
            object status;
            try
            {
                bool inShow = Globals.ThisAddIn.Application.SlideShowWindows.Count > 0;
                status = new
                {
                    isInSlideshow = inShow,
                    currentSlide = inShow ?
                        Globals.ThisAddIn.Application.SlideShowWindows[1].View.Slide.SlideIndex : 0,
                    totalSlides = inShow ?
                        Globals.ThisAddIn.Application.SlideShowWindows[1].View.Slide.Parent.Slides.Count : 0
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"读取放映状态失败(换页瞬态，返回安全值): {ex.Message}");
                status = new { isInSlideshow = false, currentSlide = 0, totalSlides = 0 };
            }

            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(status);
                var buffer = System.Text.Encoding.UTF8.GetBytes(json);

                response.ContentLength64 = buffer.Length;
                response.ContentType = "application/json; charset=utf-8";
                response.OutputStream.Write(buffer, 0, buffer.Length);
            }
            catch { }
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

        /// <summary>
        /// 控制台访问地址。已包含本次会话的访问码，直接打开/扫码即可通过鉴权。
        /// </summary>
        public static string GetConsoleUrl()
        {
            string baseUrl = $"http://{GetLocalIPAddress()}:{_serverPort}";
            return string.IsNullOrEmpty(AccessCode)
                ? baseUrl
                : $"{baseUrl}/?code={AccessCode}";
        }

        public static bool IsConsoleEnabled()
        {
            return _webConsoleEnabled;
        }

        /// <summary>
        /// 获取用于手机扫码访问的局域网 IPv4 地址。
        ///
        /// 修复：原实现取“第一个非 169.254 的 IPv4”，很容易选中 VirtualBox / Hyper-V /
        /// VMware / VPN 虚拟网卡的地址，导致生成的二维码扫出来连不上。
        /// 现改为遍历本机所有活动网卡：优先选择“有默认网关的物理网卡”，
        /// 排除虚拟网卡与常见虚拟网卡前缀，最后才回退到任意 IPv4。
        /// </summary>
        private static string GetLocalIPAddress()
        {
            // 常见的虚拟/容器网卡描述或名称关键字
            string[] virtualHints =
            {
                "virtual", "vmware", "virtualbox", "hyper-v", "hyperv",
                "loopback", "tap-", "docker", "wsl", "vpn", "ras ",
                "pseudo", "isatap", "teredo", "tunnel", "bluetooth"
            };

            try
            {
                var candidates = new List<NetworkInterfaceInfo>();

                foreach (var ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != OperationalStatus.Up)
                        continue;
                    if (ni.NetworkInterfaceType == NetworkInterfaceType.Loopback)
                        continue;

                    var ipProps = ni.GetIPProperties();
                    bool hasGateway = ipProps.GatewayAddresses.Count > 0;

                    foreach (var addr in ipProps.UnicastAddresses)
                    {
                        if (addr.Address.AddressFamily != AddressFamily.InterNetwork)
                            continue;

                        string ip = addr.Address.ToString();
                        if (ip.StartsWith("169.254.")) // APIPA，无实际路由
                            continue;

                        // VMware / VirtualBox 等虚拟网卡的典型网段
                        if (ip.StartsWith("192.168.56.") || ip.StartsWith("192.168.137."))
                            continue;

                        string desc = (ni.Description + " " + ni.Name).ToLowerInvariant();
                        bool isVirtual = false;
                        foreach (var hint in virtualHints)
                        {
                            if (desc.Contains(hint)) { isVirtual = true; break; }
                        }

                        candidates.Add(new NetworkInterfaceInfo
                        {
                            Address = ip,
                            HasGateway = hasGateway,
                            IsVirtual = isVirtual,
                            IsWireless = ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211
                        });
                    }
                }

                // 优先级：有网关的物理网卡 > 无网关的物理网卡 > 无线网卡 > 虚拟网卡
                var ordered = candidates
                    .OrderBy(c => c.IsVirtual ? 1 : 0)
                    .ThenBy(c => c.HasGateway ? 0 : 1)
                    .ToList();

                if (ordered.Count > 0)
                    return ordered[0].Address;
            }
            catch
            {
                // 获取失败则回退
            }
            return "127.0.0.1";
        }

        private class NetworkInterfaceInfo
        {
            public string Address;
            public bool HasGateway;
            public bool IsVirtual;
            public bool IsWireless;
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
                // WebSocket 握手前同样校验访问码：
                // 屏幕共享端口会把整屏画面持续推给对端，文本输入端口可以模拟键盘，
                // 两者都必须鉴权，否则等于把键盘和屏幕暴露给局域网/公网。
                if (!IsAuthorized(context.Request))
                {
                    context.Response.StatusCode = 401;
                    context.Response.Close();
                    return;
                }

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

                        // 修复空引用竞争：原实现判断 _screenWebSocket 非空后，
                        // 下一行再次读取该静态字段。两次读取之间它可能被消息处理线程
                        // dispose 并置为 null，导致偶发 NullReferenceException。
                        // 现先取到局部变量快照，整段逻辑只使用该快照。
                        var socket = _screenWebSocket;
                        if (socket != null && socket.State == WebSocketState.Open)
                        {
                            using (var ms = new MemoryStream())
                            {
                                screen.Save(ms, ImageFormat.Jpeg);
                                var buffer = ms.ToArray();
                                try
                                {
                                    socket.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Binary, true, CancellationToken.None).Wait();
                                }
                                catch (Exception sendEx)
                                {
                                    Debug.WriteLine($"发送屏幕帧失败(连接可能已断开): {sendEx.Message}");
                                }
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