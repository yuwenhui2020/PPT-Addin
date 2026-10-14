@echo off
chcp 65001 >nul
title PPT网页插件 - WEB控制台端口放行工具
color 0A

REM =========================================================================
REM  PPT网页插件 - WEB控制台端口放行脚本
REM
REM  用途：在首次使用“WEB 控制台”之前，预先完成两项需要管理员权限的初始化，
REM        之后用普通权限启动 PowerPoint 也能直接开启控制台。
REM
REM      1) HTTP URL 保留项 (urlacl)
REM         插件用 HttpListener 绑定 http://+:端口/ ，
REM         没有预先保留会报“拒绝访问 / Access is denied”。
REM
REM      2) Windows 防火墙入站放行
REM         手机等外部设备需要访问时才必需（仅本机浏览器访问可不要）。
REM         规则名与程序内 FirewallHelper 保持一致：RemoteControl_<端口>
REM
REM  端口一览（与 WebControl.cs 中定义一致）：
REM      8888  WEB 控制台页面
REM      8887  屏幕共享 (WebSocket)
REM      8889  文本/控制输入 (WebSocket)
REM
REM  用法：
REM      右键本文件 -> “以管理员身份运行”
REM      清理模式：InitWebPorts.cmd remove
REM =========================================================================

if /i "%~1"=="remove" goto :DoRemove

echo ==================================================
echo    PPT网页插件 - WEB控制台端口放行工具
echo ==================================================
echo.
echo   将完成以下初始化（端口 8888 / 8887 / 8889）：
echo     1. HTTP URL 保留项 urlacl
echo     2. Windows 防火墙入站放行
echo.
echo ==================================================
echo.

REM ---------------- 管理员权限检查与自动提升 ----------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] 检测到当前不是管理员权限，正在请求提升...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

setlocal
set PORTS=8888 8887 8889
set OK_ALL=1

echo [步骤 1/3] 配置 HTTP URL 保留项 urlacl...
for %%P in (%PORTS%) do call :AddUrlAcl %%P

echo.
echo [步骤 2/3] 配置 Windows 防火墙入站规则...
for %%P in (%PORTS%) do call :AddFirewallRule %%P

echo.
echo [步骤 3/3] 校验结果...
for %%P in (%PORTS%) do call :ShowResult %%P

echo.
echo ==================================================
if "%OK_ALL%"=="1" (
    echo 全部配置成功！
    echo 现在可以直接在插件面板点击“开启控制台”，
    echo 无需再用管理员身份运行 PowerPoint。
) else (
    echo 存在未成功的项目（上方标记为“失败”）。
    echo 常见原因：安全软件拦截、系统精简导致 netsh 功能不全。
)
echo ==================================================
echo.
pause
exit /b 0


REM =========================================================================
REM  子程序：添加 HTTP URL 保留项
REM  依次尝试 用户名 / 域\用户名 / Everyone，任一成功即可
REM =========================================================================
:AddUrlAcl
set P=%1
netsh http delete urlacl url=http://+:%P%/ >nul 2>&1

netsh http add urlacl url=http://+:%P%/ user="%USERNAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo   [OK] 端口 %P% URL保留项（用户：%USERNAME%）
    goto :eof
)

netsh http add urlacl url=http://+:%P%/ user="%USERDOMAIN%\%USERNAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo   [OK] 端口 %P% URL保留项（用户：%USERDOMAIN%\%USERNAME%）
    goto :eof
)

netsh http add urlacl url=http://+:%P%/ user="Everyone" >nul 2>&1
if %errorlevel% equ 0 (
    echo   [OK] 端口 %P% URL保留项（Everyone）
    goto :eof
)

echo   [失败] 端口 %P% URL保留项
set OK_ALL=0
goto :eof


REM =========================================================================
REM  子程序：添加防火墙入站规则
REM  规则名必须与 FirewallHelper.cs 中的 RemoteControl_端口 完全一致，
REM  这样程序运行时检测到规则已存在就不会重复添加。
REM =========================================================================
:AddFirewallRule
set P=%1
set RULE=RemoteControl_%P%

netsh advfirewall firewall show rule name="%RULE%" >nul 2>&1
if %errorlevel% equ 0 (
    echo   [OK] 端口 %P% 防火墙规则已存在（%RULE%）
    goto :eof
)

netsh advfirewall firewall add rule name="%RULE%" dir=in action=allow protocol=TCP localport=%P% >nul 2>&1
if %errorlevel% equ 0 (
    echo   [OK] 端口 %P% 防火墙入站已放行（%RULE%）
    goto :eof
)

echo   [失败] 端口 %P% 防火墙规则
set OK_ALL=0
goto :eof


REM =========================================================================
REM  子程序：显示单个端口的当前状态
REM =========================================================================
:ShowResult
set P=%1
echo   ---- 端口 %P% ----

netsh http show urlacl url=http://+:%P%/ 2>nul | findstr /i "http://+:%P%/" >nul 2>&1
if %errorlevel% equ 0 (
    echo        URL保留项：已配置
) else (
    echo        URL保留项：缺失
)

netsh advfirewall firewall show rule name="RemoteControl_%P%" 2>nul | findstr /i "RemoteControl_%P%" >nul 2>&1
if %errorlevel% equ 0 (
    echo        防火墙入站：已放行
) else (
    echo        防火墙入站：未放行
)
goto :eof


REM =========================================================================
REM  清理模式：移除 URL 保留项与防火墙规则
REM  用法：InitWebPorts.cmd remove
REM =========================================================================
:DoRemove
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] 检测到当前不是管理员权限，正在请求提升...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs -ArgumentList 'remove'"
    exit /b
)

echo ==================================================
echo    正在清理端口放行配置...
echo ==================================================
echo.

for %%P in (8888 8887 8889) do call :RemoveOne %%P

echo.
echo 清理完成。之后开启 WEB 控制台可能需要重新以管理员
echo 身份运行 PowerPoint，或再次运行本脚本重新放行。
echo.
pause
exit /b 0

:RemoveOne
set P=%1
netsh http delete urlacl url=http://+:%P%/ >nul 2>&1
echo   端口 %P%：URL保留项已删除
netsh advfirewall firewall delete rule name="RemoteControl_%P%" >nul 2>&1
echo   端口 %P%：防火墙规则已删除
goto :eof
