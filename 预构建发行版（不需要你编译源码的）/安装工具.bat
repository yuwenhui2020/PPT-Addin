@echo off
chcp 936 >nul
title PowerPoint 网页插件 安装程序
color 0A

:: =========================================================================
::  提权后通过参数直接进入对应流程（避免再次显示菜单丢失选择）
:: =========================================================================
if '%1'=='install' goto install
if '%1'=='uninstall' goto uninstall
if '%1'=='repair' goto repair

:menu
cls
echo ========================================
echo    PowerPoint 网页插件 安装程序
echo ========================================
echo.
echo   [1] 安装插件（含运行库自检与静默安装）
echo   [2] 卸载插件（保留系统运行库，不卸载）
echo   [3] 修复插件（先卸载再重装 + 运行库检测）
echo   [4] 退出
echo.
echo ========================================

set /p choice=请选择操作 (1-4): 

if "%choice%"=="1" call :NeedAdmin install
if "%choice%"=="2" call :NeedAdmin uninstall
if "%choice%"=="3" call :NeedAdmin repair
if "%choice%"=="4" exit /b
goto menu


:: =========================================================================
::  管理员权限自检：无权限则提权后携带操作参数重启自身
:: =========================================================================
:NeedAdmin
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' neq '0' (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs -ArgumentList '%1'"
    exit /b
)
if '%1'=='install' goto install
if '%1'=='uninstall' goto uninstall
if '%1'=='repair' goto repair
goto menu


:: =========================================================================
::  安装（仅安装需要进行运行库检测）
:: =========================================================================
:install
cls
echo ========================================
echo        安装 PowerPoint 网页插件
echo ========================================
echo.
echo [第一步] 运行库自检与静默安装（仅安装/修复需要）
echo.
call :RuntimeEnsure
echo.
echo [第二步] 注册并安装插件...
echo.
call :PluginRegister
echo.
echo ========================================
echo  安装完成！
echo ========================================
echo  请重启 PowerPoint（或 WPS）以使插件生效。
echo  若运行库安装提示需要重启，请重启电脑后再使用。
echo.
echo 按下任意键返回菜单... &pause>nul
goto menu


:: =========================================================================
::  卸载（不卸载系统运行库）
:: =========================================================================
:uninstall
cls
echo ========================================
echo        卸载 PowerPoint 网页插件
echo ========================================
echo.
echo 警告：此操作将删除插件注册信息（不会卸载系统运行库）。
echo.
set /p confirm=确认卸载？(Y/N): 
if /i "%confirm%" neq "Y" goto menu
echo.
call :PluginRemove
echo.
echo ========================================
echo  卸载完成！系统运行库保持不变。
echo ========================================
echo 按下任意键返回菜单... &pause>nul
goto menu


:: =========================================================================
::  修复 = 卸载后重新安装（仍需检测运行库是否安装成功）
:: =========================================================================
:repair
cls
echo ========================================
echo        修复 PowerPoint 网页插件
echo ========================================
echo.
echo 修复流程：先卸载现有插件，再重新安装并检测运行库。
echo.
echo [第一步] 卸载现有插件（保留运行库）...
echo.
call :PluginRemove
echo.
echo [第二步] 检测并安装运行库（确认安装成功）...
echo.
call :RuntimeEnsure
echo.
echo [第三步] 重新注册并安装插件...
echo.
call :PluginRegister
echo.
echo ========================================
echo  修复完成！
echo ========================================
echo  请重启 PowerPoint（或 WPS）以使插件生效。
echo.
echo 按下任意键返回菜单... &pause>nul
goto menu


:: =========================================================================
::  运行库自检 + 静默安装 + 安装后验证
::   errorlevel 0=全部就绪 1=存在未成功安装项
::   仅安装与修复调用本例程；卸载不调用。
:: =========================================================================
:RuntimeEnsure
set "RUNTIME_FAIL=0"

:: ---------- 定位运行库目录 ----------
set "RUNTIMEDIR="
if exist "%~dp0runtime\" set "RUNTIMEDIR=%~dp0runtime\"
if not defined RUNTIMEDIR if exist "%~dp0..\bin\runtime\" set "RUNTIMEDIR=%~dp0..\bin\runtime\"
if not defined RUNTIMEDIR if exist "%~dp0..\..\bin\runtime\" set "RUNTIMEDIR=%~dp0..\..\bin\runtime\"
if not defined RUNTIMEDIR (
    echo [!] 未找到运行库目录 runtime\。
    echo     请将本安装程序与 runtime\ 文件夹放在一起（runtime 内含三个运行库安装程序），
    echo     或把 bin\runtime 复制到本安装程序同目录下。
    echo     已跳过运行库安装，仅安装插件本体。
    echo.
    pause
    exit /b 1
)

echo 检测必备运行库，缺失将自动静默安装（显示进度）...
echo.

echo [1/3] .NET Framework 4.8 运行库
call :CheckNetFx48
if errorlevel 1 (
    echo     未安装，开始静默安装...
    call :InstallRuntime "%RUNTIMEDIR%NDP48-x86-x64-AllOS-ENU.exe" ".NET Framework 4.8" "/q /norestart"
    call :CheckNetFx48
    if errorlevel 1 ( echo     [!] 安装后仍未检测到，请手动运行 runtime\ 中的安装程序 & set "RUNTIME_FAIL=1" ) else ( echo     [OK] 已成功安装并验证 )
) else (
    echo     [OK] 已安装，跳过
)

echo [2/3] WebView2 运行时（网页显示组件）
call :CheckWebView2
if errorlevel 1 (
    echo     未安装，开始静默安装...
    call :InstallRuntime "%RUNTIMEDIR%MicrosoftEdgeWebView2RuntimeInstallerX64.exe" "WebView2 运行时" "/silent /install"
    call :CheckWebView2
    if errorlevel 1 ( echo     [!] 安装后仍未检测到，请手动运行 runtime\ 中的安装程序 & set "RUNTIME_FAIL=1" ) else ( echo     [OK] 已成功安装并验证 )
) else (
    echo     [OK] 已安装，跳过
)

echo [3/3] .NET 8 桌面运行库
call :CheckDotNet8Desktop
if errorlevel 1 (
    echo     未安装，开始静默安装...
    call :InstallRuntime "%RUNTIMEDIR%windowsdesktop-runtime-8.0.16-win-x64.exe" ".NET 8 桌面运行库" "/install /quiet /norestart"
    call :CheckDotNet8Desktop
    if errorlevel 1 ( echo     [!] 安装后仍未检测到，请手动运行 runtime\ 中的安装程序 & set "RUNTIME_FAIL=1" ) else ( echo     [OK] 已成功安装并验证 )
) else (
    echo     [OK] 已安装，跳过
)

echo.
if "%RUNTIME_FAIL%"=="1" ( echo [运行库自检] 完成，但存在未成功安装的运行库，请检查上方提示。 ) else ( echo [运行库自检] 完成，全部运行库就绪。 )
echo.
exit /b %RUNTIME_FAIL%


:: =========================================================================
::  插件注册（HKCU / HKLM / WPS）+ 端口预初始化
:: =========================================================================
:PluginRegister
:: ---------- 检查 VSTO 文件 ----------
if exist "%~dp0PowerPointAddIn.vsto" (
    echo [OK] 找到 PowerPointAddIn.vsto
) else (
    echo [X] 错误: 未找到 PowerPointAddIn.vsto，请确保它与本安装程序在同一目录。
    echo 按下任意键返回菜单... &pause>nul
    goto menu
)

set "REG_PATH=HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn"
set "MANIFEST=file:///%~dp0PowerPointAddIn.vsto|vstolocal"

echo.
echo [插件注册] 写入注册表...
reg add "%REG_PATH%" /f >nul 2>&1 && echo   [OK] HKCU-PPT  || echo   [X] HKCU-PPT 创建失败
reg add "%REG_PATH%" /v "FriendlyName" /t REG_SZ /d "网页插件" /f >nul 2>&1 && echo   [OK] HKCU-FriendlyName  || echo   [X] 失败
reg add "%REG_PATH%" /v "Description" /t REG_SZ /d "一款全能的Office网页插件" /f >nul 2>&1 && echo   [OK] HKCU-Description || echo   [X] 失败
reg add "%REG_PATH%" /v "LoadBehavior" /t REG_DWORD /d "3" /f >nul 2>&1 && echo   [OK] HKCU-LoadBehavior || echo   [X] 失败
reg add "%REG_PATH%" /v "CommandLineSafe" /t REG_DWORD /d "1" /f >nul 2>&1 && echo   [OK] HKCU-CommandLineSafe || echo   [X] 失败
reg add "%REG_PATH%" /v "Manifest" /t REG_SZ /d "%MANIFEST:\=\\%" /f >nul 2>&1 && echo   [OK] HKCU-Manifest || echo   [X] 失败

(
    echo Windows Registry Editor Version 5.00
    echo.
    echo [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn]
    echo "FriendlyName"="网页插件"
    echo "Description"="一款全能的Office网页插件"
    echo "LoadBehavior"=dword:00000003
    echo "CommandLineSafe"=dword:00000001
    echo "Manifest"="%MANIFEST:\=\\%"
) > "%TEMP%\ppt_system.reg"
reg import "%TEMP%\ppt_system.reg" >nul 2>&1 && echo   [OK] HKLM-PPT-64位兼容 || echo   [!] HKLM-PPT-64位兼容失败（可能需更高权限）

(
    echo Windows Registry Editor Version 5.00
    echo.
    echo [HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn]
    echo "FriendlyName"="网页插件"
    echo "Description"="一款全能的Office网页插件"
    echo "LoadBehavior"=dword:00000003
    echo "CommandLineSafe"=dword:00000001
    echo "Manifest"="%MANIFEST:\=\\%"
) > "%TEMP%\ppt_wow64.reg"
reg import "%TEMP%\ppt_wow64.reg" >nul 2>&1 && echo   [OK] HKLM-PPT-32位兼容 || echo   [!] HKLM-PPT-32位兼容失败

reg add "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL" /v "PowerPointAddIn" /t REG_SZ /d "" /f >nul 2>&1
if errorlevel 1 ( echo   [!] WPS 注册失败 ) else ( echo   [OK] WPS Office 兼容 )

:: ---------- WEB 控制台端口预初始化（8887-8892） ----------
echo.
echo [端口初始化] 预配置 WEB 控制台端口（urlacl + 防火墙），之后开启控制台无需 UAC...
for %%P in (8887 8888 8889 8890 8891 8892) do call :PortInit %%P
echo   [OK] 端口初始化完毕

del "%TEMP%\ppt_system.reg" >nul 2>&1
del "%TEMP%\ppt_wow64.reg" >nul 2>&1
exit /b


:: =========================================================================
::  插件卸载（仅删除注册信息，不触碰系统运行库）
:: =========================================================================
:PluginRemove
echo 正在卸载插件（保留系统运行库）...
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f>nul 2>&1 && echo   [OK] HKCU-PPT 删除成功    || echo   [X] HKCU-PPT 删除失败
reg delete "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL" /f>nul 2>&1 && echo   [OK] HKCU-WPS 兼容删除成功    || echo   [X] HKCU-WPS 兼容删除失败
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f>nul 2>&1 && echo   [OK] HKLM-PPT-64 删除成功    || echo   [X] HKLM-PPT-64 删除失败
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f>nul 2>&1 && echo   [OK] HKLM-PPT-32 删除成功    || echo   [X] HKLM-PPT-32 删除失败
"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Uninstall file:///%~dp0PowerPointAddIn.vsto /s >nul 2>&1 && echo   [OK] APP-PPT 删除成功    || echo   [X] APP-PPT 删除失败（可能未曾通过双击 vsto 安装）
exit /b


:: =========================================================================
::  运行库检测子程序：errorlevel 0=已安装 1=未安装
:: =========================================================================
:CheckNetFx48
powershell -NoProfile -Command "$r=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue; if($r -and $r.Release -ge 528040){exit 0}else{exit 1}"
exit /b

:CheckWebView2
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "Microsoft Edge WebView2 Runtime" >nul
if not errorlevel 1 exit /b 0
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "Microsoft Edge WebView2 Runtime" >nul
if not errorlevel 1 exit /b 0
exit /b 1

:CheckDotNet8Desktop
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "Microsoft Windows Desktop Runtime" >nul
if not errorlevel 1 exit /b 0
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "Microsoft Windows Desktop Runtime" >nul
if not errorlevel 1 exit /b 0
exit /b 1


:: =========================================================================
::  运行库静默安装子程序（带进度显示）
::   %1 = 安装程序完整路径
::   %2 = 显示名称
::   %3 = 静默安装参数
:: =========================================================================
:InstallRuntime
set "EXE=%~1"
set "NAME=%~2"
set "ARGS=%~3"
if not exist "%EXE%" (
    echo     [!] 未找到安装程序：%EXE%
    echo         请确认 runtime\ 目录完整（含三个运行库安装程序）。
    exit /b 1
)
echo.
echo   -- 正在安装 %NAME% ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$p=Start-Process -FilePath '%EXE%' -ArgumentList '%ARGS%' -PassThru; $pct=0; while(-not $p.HasExited){ $pct=[math]::Min(99,$pct+2); Write-Progress -Activity '安装 %NAME%' -Status '正在安装，请稍候...' -PercentComplete $pct; Start-Sleep -Milliseconds 300 }; Write-Progress -Activity '安装 %NAME%' -Completed; Write-Host ('    完成（退出码 '+$p.ExitCode+'）'); if($p.ExitCode -eq 3010 -or $p.ExitCode -eq 1641){ Write-Host '    [提示] 该运行库需要重启系统后生效' }"
exit /b 0


:: =========================================================================
::  子程序：初始化单个 WEB 控制台端口
::   1) urlacl：HttpListener 绑定 http://+:端口/ 必须先保留，否则拒绝访问
::   2) 防火墙：规则名须与 FirewallHelper/WebControl 的 RemoteControl_端口 一致
::  本安装程序已具备管理员权限，无需再次提权；且操作幂等，可重复执行。
::  端口范围 8887-8892 覆盖“默认三段”与“8888 被占用后的后移段”，使首次开启控制台 0 次 UAC。
:: =========================================================================
:PortInit
set PP=%1

netsh http delete urlacl url=http://+:%PP%/ >nul 2>&1
netsh http add urlacl url=http://+:%PP%/ user="%USERNAME%" >nul 2>&1
if %errorlevel% equ 0 goto :PortFw
netsh http add urlacl url=http://+:%PP%/ user="%USERDOMAIN%\%USERNAME%" >nul 2>&1
if %errorlevel% equ 0 goto :PortFw
netsh http add urlacl url=http://+:%PP%/ user="Everyone" >nul 2>&1

:PortFw
netsh advfirewall firewall show rule name="RemoteControl_%PP%" >nul 2>&1
if %errorlevel% equ 0 (
    echo     [OK] 端口 %PP% 已放行
    goto :eof
)
netsh advfirewall firewall add rule name="RemoteControl_%PP%" dir=in action=allow protocol=TCP localport=%PP% >nul 2>&1
if %errorlevel% equ 0 (
    echo     [OK] 端口 %PP% 已放行
) else (
    echo     [!] 端口 %PP% 防火墙放行失败（不影响本机浏览器访问）
)
goto :eof
