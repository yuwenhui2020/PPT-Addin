using System;
using System.Diagnostics;

namespace PowerPointAddIn
{
    public static class FirewallHelper
    {
        public static void AddFirewallRule(int port)
        {
            try
            {
                // 检查是否已存在规则
                if (CheckFirewallRuleExists(port))
                    return;

                // 添加入站规则
                AddInboundRule(port);

                Console.WriteLine($"已添加防火墙规则，允许端口 {port} 的TCP连接");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加防火墙规则失败: {ex.Message}");
            }
        }

        private static bool CheckFirewallRuleExists(int port)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "netsh",
                    Arguments = $"advfirewall firewall show rule name=\"RemoteControl_{port}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            return output.Contains("RemoteControl_" + port);
        }

        private static void AddInboundRule(int port)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "netsh",
                    Arguments = $"advfirewall firewall add rule name=\"RemoteControl_{port}\" dir=in action=allow protocol=TCP localport={port}",
                    Verb = "runas", // 以管理员权限运行
                    UseShellExecute = true,
                    CreateNoWindow = true
                }
            };

            process.Start();
            process.WaitForExit();
        }

        public static void RemoveFirewallRule(int port)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "netsh",
                        Arguments = $"advfirewall firewall delete rule name=\"RemoteControl_{port}\"",
                        Verb = "runas",
                        UseShellExecute = true,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                process.WaitForExit();

                Console.WriteLine($"已移除防火墙规则，端口 {port}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移除防火墙规则失败: {ex.Message}");
            }
        }
    }
}