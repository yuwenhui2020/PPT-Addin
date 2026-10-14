using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Generic;

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

        // 日志写入的“安全包装”：诊断绝不能影响正常功能，任何失败都静默忽略
        private static void WriteTraceSafe(string message)
        {
            try { PowerPointAddIn.ThisAddIn.WriteTrace(message); } catch { }
        }

        // 记录当前覆盖层所依附的放映窗口句柄，便于翻页/定时刷新时重新定位，
        // 以及检测该窗口是否仍存在（演讲者视图关闭或窗口重建时句柄会失效）。
        public IntPtr OwnerHwnd { get; set; }

        // 上一次“有效定位”的结果，用于在换页过渡、窗口几何暂不可得时，
        // 仍能停留在正确位置并可见（避免退回屏幕尺寸/1px 而闪退到屏幕外导致“看不见”）。
        public int LastGoodLeft, LastGoodTop, LastGoodWidth, LastGoodHeight;
        public bool HasLastGood = false;

        // 把 localhost 归一化为 127.0.0.1：
        // Windows 上 localhost 可能优先解析到 IPv6 的 ::1，而很多本地教学/开发服务器
        // 只监听 IPv4(127.0.0.1)。导航主文档时会因“::1 失败→回退 127.0.0.1”而成功，
        // 但随后页面发起的大量“懒加载分块(动态 import 的 .async.js / chunk.css / 图片)”
        // 并发请求，容易落在 ::1 这条失败链路上、连接被重置，Chromium 将其归为
        // ERR_CACHE_READ_FAILURE —— 表现为“主包能加载、分块资源加载失败”，正常浏览器却正常。
        // 强制走 IPv4(127.0.0.1) 可彻底消除该双栈竞态。
        private static string NormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return url;
            try
            {
                var uri = new Uri(url);
                if (uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase))
                {
                    var builder = new UriBuilder(uri) { Host = "127.0.0.1" };
                    return builder.ToString();
                }
            }
            catch { }
            return url;
        }

        /// <summary>
        /// 判断“将要导航的目标 URL”是否与当前已加载的 URL 等价（按归一化后的字符串比较）。
        /// 关键：CurrentUrl 保存的是 NormalizeUrl 后的地址（如 localhost 已改写为 127.0.0.1），
        /// 而调用方传入的可能是未归一化的原始地址。若直接用原始地址与 CurrentUrl 比较，
        /// 两者永远不相等，导致每次显示都重新导航 → 反复重载白屏（PowerPoint 浪费带宽，
        /// WPS 放映时更会因监控定时器高频触发而形成闪白屏循环）。故统一归一化后再比较。
        /// </summary>
        public bool IsCurrentUrl(string url)
        {
            return string.Equals(CurrentUrl, NormalizeUrl(url), StringComparison.OrdinalIgnoreCase);
        }

        public void NavigateToUrl(string url)
        {
            try
            {
                url = NormalizeUrl(url);
                CurrentUrl = url;
                if (webView21.CoreWebView2 != null)
                {
                    webView21.Source = new Uri(url);
                }
                else
                {
                    // 初始化尚未完成：记录待导航地址，InitializeAsync 完成后会补导航
                    _pendingUrl = url;
                }
            }
            catch (Exception ex)
            {
                PowerPointAddIn.ThisAddIn.WriteTrace($"[WebView2] 导航失败: {ex.Message} Url={url}");
                MessageBox.Show($"导航失败: {ex.Message}");
            }
        }

        public void AddWebView(string url, int width, int height)
        {
            try
            {
                url = NormalizeUrl(url);
                CurrentUrl = url;

                // 本地文件：记录文件与同目录资源是否存在，便于诊断“资源加载失败”
                if (url.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        string local = new Uri(url).LocalPath;
                        string dir = Path.GetDirectoryName(local);
                        WriteTraceSafe($"[添加网页] 本地文件={local} 存在={File.Exists(local)}");
                        if (File.Exists(local))
                        {
                            WriteTraceSafe($"[添加网页] 文件大小={new FileInfo(local).Length} 字节");
                            // 列出同目录文件，用于确认 css/js/图片 是否随 HTML 一起存在
                            try
                            {
                                var siblings = Directory.GetFiles(dir).Take(30).Select(Path.GetFileName);
                                WriteTraceSafe("[添加网页] 同目录文件: " + string.Join(", ", siblings));
                            }
                            catch { }
                        }
                        else
                        {
                            WriteTraceSafe($"[添加网页] 目录={dir} 目录存在={Directory.Exists(dir)}");
                        }
                    }
                    catch { }
                }
                else
                {
                    WriteTraceSafe($"[添加网页] 远程地址={url}");
                }

                if (webView21.CoreWebView2 != null)
                {
                    webView21.Source = new Uri(url);
                }
                else
                {
                    // 初始化尚未完成：记录待导航地址，InitializeAsync 完成后会补导航
                    _pendingUrl = url;
                }

                webView21.Width = width;
                webView21.Height = height;
                this.ClientSize = new System.Drawing.Size(width, height);
                webView21.Dock = DockStyle.Fill;
            }
            catch (Exception ex)
            {
                WriteTraceSafe($"[添加网页] 异常: {ex.Message} Url={url}");
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
        // =====================================================================
        // 关键修复：WebView2 环境（Environment）必须在整个进程内【唯一共享】。
        //
        // 原实现是每个 WebBrowserForm 各自 CreateAsync 一个环境，却都指向同一个
        // userDataFolder。WebView2 明确规定：同一 user data folder 不能被多个环境
        // 同时使用。插件会为每个窗口(全屏/演讲者视图)、每张幻灯片各建一个 Form，
        // 于是出现多个环境争抢同一份缓存目录 —— 表现为：
        //   * 某些设备上网页“资源加载失败”(css/js/图片 404 或加载中断)，而正常浏览器没问题
        //   * 时好时坏、与设备性能和时序相关（正是“某些设备上可能会”的原因）
        // 现改为：静态方法惰性创建单一共享环境，所有 WebView2 控件复用它。
        // =====================================================================
        private static Microsoft.Web.WebView2.Core.CoreWebView2Environment _sharedEnvironment;
        private static readonly object _environmentLock = new object();
        private static Task<Microsoft.Web.WebView2.Core.CoreWebView2Environment> _environmentTask;

        /// <summary>
        /// 获取进程内唯一共享的 WebView2 环境（线程安全、带失败重试）。
        /// </summary>
        private static Task<Microsoft.Web.WebView2.Core.CoreWebView2Environment> GetSharedEnvironmentAsync()
        {
            lock (_environmentLock)
            {
                // 已成功或正在创建中：直接复用
                if (_environmentTask != null && !_environmentTask.IsFaulted && !_environmentTask.IsCanceled)
                    return _environmentTask;

                // 上一次创建失败：清空以便后续重试
                _environmentTask = CreateEnvironmentAsync();
                return _environmentTask;
            }
        }

        private static async Task<Microsoft.Web.WebView2.Core.CoreWebView2Environment> CreateEnvironmentAsync()
        {
            string appDataPath = ResolveUserDataFolder();
            if (string.IsNullOrEmpty(appDataPath))
            {
                throw new Exception("无法创建 WebView2 用户数据目录（本地磁盘不可用或被权限阻止）。");
            }

            // 修复 ERR_CACHE_READ_FAILURE（日志中本地皮肤 PNG 等资源加载失败的根因）：
            // 上一故障期（多环境争抢同一 userDataFolder）在 Cache 里留下了损坏的缓存索引，
            // 由于之前改为“退出不再清空缓存”，这笔坏缓存被长期保留，导致每次命中都读取失败。
            // 这里只删除 HTTP 磁盘缓存子目录(Cache/GPUCache)，不删整个 userDataFolder，
            // 既清掉损坏索引，又保留 profile/cookie，且避免在浏览器运行期删整目录带来的风险。
            try
            {
                foreach (string sub in new[] { "Cache", "GPUCache", "Code Cache" })
                {
                    string subDir = Path.Combine(appDataPath, sub);
                    if (Directory.Exists(subDir))
                    {
                        Directory.Delete(subDir, true);
                        PowerPointAddIn.ThisAddIn.WriteTrace($"[环境] 已清理旧的磁盘缓存目录: {subDir}");
                    }
                }
            }
            catch (Exception cacheEx)
            {
                PowerPointAddIn.ThisAddIn.WriteTrace($"[环境] 清理旧缓存失败(忽略，不影响使用): {cacheEx.Message}");
            }

            Directory.CreateDirectory(appDataPath);

            // 权限设置仅作“尽力而为”：失败不影响后续创建（曾因 ACL 异常导致整窗白屏）。
            try
            {
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
            }
            catch (Exception aclEx)
            {
                PowerPointAddIn.ThisAddIn.WriteTrace($"[WebView2] 用户数据目录ACL设置失败(忽略，不影响使用): {aclEx.Message}");
            }

            // 与历史一致的运行参数（本地文件访问 / 关闭磁盘缓存 / 强制 HTTP/1.1）。
            // --disable-web-security：本地 file:// 页面通过 fetch/XHR/ES Module/Worker
            //   引用同目录资源时的同源策略拦截；--disable-cache：杜绝磁盘缓存索引损坏；
            // --disable-http2：强制 HTTP/1.1，避免某些本地服务器 HTTP/2 多路复用缺陷
            //   导致懒加载分块被截断（表现为 ERR_CACHE_READ_FAILURE）。
            string args = "--disable-web-security --allow-file-access-from-files --disable-cache --disable-http2";

            // 可选：若随程序附带了“固定版本”运行时（详见 ResolveBrowserExecutableFolder），
            // 优先使用它，彻底摆脱对目标电脑系统运行时的依赖。
            string browserFolder = ResolveBrowserExecutableFolder();

            // 预检：确认能找到可用的 WebView2 运行时（系统或附带目录）。
            // 很多“装了也起不来”的机器，其实是系统运行时缺失 / 被组策略禁用 / 版本过低，
            // 这里提前明确报错，避免静默白屏。
            string runtimeVer = null;
            try
            {
                runtimeVer = Microsoft.Web.WebView2.Core.CoreWebView2Environment
                    .GetAvailableBrowserVersionString(browserFolder);
            }
            catch (Exception verEx)
            {
                PowerPointAddIn.ThisAddIn.WriteTrace($"[环境] 查询可用运行时版本失败: {verEx.Message}");
            }
            if (string.IsNullOrWhiteSpace(runtimeVer))
            {
                throw new Exception(browserFolder == null
                    ? "未检测到系统 WebView2 运行时（GetAvailableBrowserVersionString 返回空）。请安装 WebView2 运行时，或随程序附带固定版本运行时目录 WebView2Runtime\\。"
                    : "随附的固定版本 WebView2 运行时不完整（缺少 WebView2Loader.dll / msedgewebview2.exe）。");
            }

            PowerPointAddIn.ThisAddIn.WriteTrace($"[环境] 开始创建 WebView2 环境, 数据目录={appDataPath}, 运行时={runtimeVer}, 参数={args}");

            var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                userDataFolder: appDataPath,
                browserExecutableFolder: browserFolder,
                options: new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions(args));

            PowerPointAddIn.ThisAddIn.WriteTrace($"[环境] 创建成功, 浏览器版本={env?.BrowserVersionString}");
            return env;
        }

        /// <summary>
        /// 解析 WebView2 的 user data folder。必须位于本地磁盘；域环境可能把
        /// LocalApplicationData 重定向到网络盘（文件夹重定向），用网络路径会让
        /// 初始化失败、白屏。因此按可用性做多级回退：本地 AppData -> 程序目录 -> 系统 TEMP。
        /// </summary>
        private static string ResolveUserDataFolder()
        {
            var candidates = new List<string>();
            try { candidates.Add(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PowerPointAddIn", "WebView2Data")); } catch { }
            try { candidates.Add(Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "PowerPointAddIn_WebView2Data")); } catch { }
            try { candidates.Add(Path.Combine(
                Path.GetTempPath(), "PowerPointAddIn_WebView2Data")); } catch { }

            foreach (var dir in candidates)
            {
                try
                {
                    Directory.CreateDirectory(dir);
                    string probe = Path.Combine(dir, ".__probe");
                    File.WriteAllText(probe, "1");
                    File.Delete(probe);
                    return dir;
                }
                catch { }
            }
            return null;
        }

        /// <summary>
        /// 解析可选的“固定版本” WebView2 运行时目录。若程序同目录存在 WebView2Runtime\
        /// （解压后的固定版本，内含 WebView2Loader.dll 或 msedgewebview2.exe），优先使用，
        /// 使程序不依赖“目标电脑已安装系统 WebView2 运行时”——这是跨设备分发时
        /// “某些电脑装了也起不来”的最可靠解法。未附带则返回 null（回退系统运行时）。
        /// </summary>
        private static string ResolveBrowserExecutableFolder()
        {
            try
            {
                string root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebView2Runtime");
                if (!Directory.Exists(root)) return null;
                var dirs = new List<string> { root };
                dirs.AddRange(Directory.GetDirectories(root));
                foreach (var d in dirs)
                {
                    if (File.Exists(Path.Combine(d, "WebView2Loader.dll")) ||
                        File.Exists(Path.Combine(d, "msedgewebview2.exe")))
                        return d;
                }
            }
            catch { }
            return null;
        }

        private static bool _initErrorShown = false;
        /// <summary>
        /// 把 WebView2 初始化失败明确提示给用户（只弹一次），并给出可操作建议，
        /// 避免“装了也白屏、却没有任何提示”的困惑。
        /// </summary>
        private static void ShowInitError(string detail)
        {
            PowerPointAddIn.ThisAddIn.WriteTrace($"[WebView2] 初始化失败(已提示用户): {detail}");
            if (_initErrorShown) return;
            _initErrorShown = true;
            try
            {
                MessageBox.Show(
                    "无法启动网页显示组件（WebView2）。\n\n" +
                    "常见原因：\n" +
                    "  1) 系统未安装 WebView2 运行时，或版本过低；\n" +
                    "  2) 缺少 Visual C++ 2015-2022 运行库（WebView2 依赖它）；\n" +
                    "  3) 被组策略禁用，或当前 Windows 版本过低。\n\n" +
                    "解决办法：\n" +
                    "  a) 安装 WebView2 运行时：\n     https://go.microsoft.com/fwlink/p/?LinkId=2124703\n" +
                    "  b) 安装 Visual C++ 运行库（x64 与 x86 都装）。\n\n" +
                    $"技术信息：{detail}",
                    "网页显示组件初始化失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }

        private async void InitializeAsync()
        {
            try
            {
                var env = await GetSharedEnvironmentAsync();
                await webView21.EnsureCoreWebView2Async(env);

                var core = webView21.CoreWebView2;

                core.NavigationStarting += (s, e) =>
                    PowerPointAddIn.ThisAddIn.WriteTrace($"[导航开始] {e.Uri}");
                core.ContentLoading += (s, e) =>
                    PowerPointAddIn.ThisAddIn.WriteTrace("[内容加载中]");
                core.DOMContentLoaded += (s, e) =>
                    PowerPointAddIn.ThisAddIn.WriteTrace("[DOM就绪]");
                core.NavigationCompleted += webView21_NavigationCompleted;
                core.NewWindowRequested += CoreWebView2_NewWindowRequested;
                core.WebResourceResponseReceived += CoreWebView2_WebResourceResponseReceived;

                // 进程崩溃/无响应：“资源加载失败”的常见根因之一，必须捕获
                core.ProcessFailed += (s, e) =>
                {
                    PowerPointAddIn.ThisAddIn.WriteTrace(
                        $"[进程失败] 类型={e.ProcessFailedKind} 原因={e.Reason} 退出码={e.ExitCode} 进程={e.ProcessDescription}");
                    // 浏览器进程异常退出多为缺少 VC++ 运行库或被杀软拦截：提示一次
                    if (e.ProcessFailedKind == Microsoft.Web.WebView2.Core.CoreWebView2ProcessFailedKind.BrowserProcessExited)
                    {
                        ShowInitError($"WebView2 浏览器进程异常（{e.Reason}）。多为缺少 Visual C++ 运行库或被安全软件拦截。");
                    }
                };

                // 用 CDP 捕获浏览器控制台日志与网络失败
                // （浏览器里看到的 “Failed to load resource / CORS” 就来自这些事件）
                try
                {
                    await core.CallDevToolsProtocolMethodAsync("Runtime.enable", "{}");
                    await core.CallDevToolsProtocolMethodAsync("Log.enable", "{}");
                    await core.CallDevToolsProtocolMethodAsync("Network.enable", "{}");
                    PowerPointAddIn.ThisAddIn.WriteTrace("[诊断] CDP已启用(Runtime/Log/Network)");

                    var logRecv = core.GetDevToolsProtocolEventReceiver("Log.entryAdded");
                    logRecv.DevToolsProtocolEventReceived += (s, e) =>
                        PowerPointAddIn.ThisAddIn.WriteTrace($"[控制台Log] {e.ParameterObjectAsJson}");

                    var excRecv = core.GetDevToolsProtocolEventReceiver("Runtime.exceptionThrown");
                    excRecv.DevToolsProtocolEventReceived += (s, e) =>
                        PowerPointAddIn.ThisAddIn.WriteTrace($"[JS异常] {e.ParameterObjectAsJson}");

                    var failRecv = core.GetDevToolsProtocolEventReceiver("Network.loadingFailed");
                    failRecv.DevToolsProtocolEventReceived += (s, e) =>
                        PowerPointAddIn.ThisAddIn.WriteTrace($"[网络加载失败] {e.ParameterObjectAsJson}");

                    // 记录 localhost 的 HTTP 响应状态码：若服务器真的返回了 200 但该资源
                    // 随后又出现在“网络加载失败”里，说明响应在传输中途被截断/连接被重置；
                    // 若根本没有“网络响应”记录，说明请求根本没到达服务器（连接层问题）。
                    // 据此可区分“服务器侧”与“传输层”根因。
                    var respRecv = core.GetDevToolsProtocolEventReceiver("Network.responseReceived");
                    respRecv.DevToolsProtocolEventReceived += (s, e) =>
                    {
                        try
                        {
                            string json = e.ParameterObjectAsJson ?? "";
                            if (json.Contains("127.0.0.1") || json.Contains("localhost"))
                                PowerPointAddIn.ThisAddIn.WriteTrace($"[网络响应] {json}");
                        }
                        catch { }
                    };
                }
                catch (Exception cdpEx)
                {
                    PowerPointAddIn.ThisAddIn.WriteTrace($"[诊断] CDP启用失败(仅影响诊断采集): {cdpEx.Message}");
                }

                PowerPointAddIn.ThisAddIn.WriteTrace("[WebView2] 控件初始化完成(共享环境)");

                // 若在初始化完成前就已经调用过 AddWebView/NavigateToUrl，此处补一次导航
                if (!string.IsNullOrEmpty(_pendingUrl))
                {
                    string pending = NormalizeUrl(_pendingUrl);
                    _pendingUrl = null;
                    PowerPointAddIn.ThisAddIn.WriteTrace($"[补导航] {pending}");
                    core.Navigate(pending);
                }
            }
            catch (Exception ex)
            {
                ShowInitError(ex?.Message ?? ex?.ToString());
            }
        }

        // 初始化完成前待导航的 URL
        private string _pendingUrl;

        /// <summary>
        /// 记录资源加载失败（4xx/5xx），便于排查“资源加载失败”的具体 URL。
        /// 日志：%LOCALAPPDATA%\PPT-Addin\webview.log
        /// </summary>
        private void CoreWebView2_WebResourceResponseReceived(
            object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebResourceResponseReceivedEventArgs e)
        {
            try
            {
                var response = e.Response;
                if (response != null && response.StatusCode >= 400)
                {
                    string reason = null;
                    try { reason = response.ReasonPhrase; } catch { }
                    PowerPointAddIn.ThisAddIn.WriteTrace(
                        $"[资源失败] Status={response.StatusCode} {reason} Url={e.Request?.Uri} 方法={e.Request?.Method}");
                }
            }
            catch { }
        }
        private void CoreWebView2_NewWindowRequested(object sender,Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs e)
        {
            e.Handled = true;
            webView21.CoreWebView2.Navigate(e.Uri);
        }
        private void webView21_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            try
            {
                if (!e.IsSuccess)
                {
                    // 记录失败详情：WebErrorStatus 能直接指明是被拦截、文件不存在还是网络问题
                    PowerPointAddIn.ThisAddIn.WriteTrace(
                        $"[WebView2] 导航失败 IsSuccess=false Status={e.WebErrorStatus} HttpStatus={e.HttpStatusCode} Url={CurrentUrl}");
                }
                else
                {
                    PowerPointAddIn.ThisAddIn.WriteTrace($"[WebView2] 导航成功 Url={CurrentUrl}");
                }
            }
            catch { }
        }
        
    }
}