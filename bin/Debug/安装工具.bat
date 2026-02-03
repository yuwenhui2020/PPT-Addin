
@echo off
chcp 65001 >nul
title PowerPoint插件管理工具
color 0A
:: 检查并获取管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' neq '0' (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs -ArgumentList 'install'"
    exit /b
)
:menu
cls
echo ========================================
echo    PowerPoint插件管理工具
echo ========================================
echo.
echo   [1]安装插件 
echo   [2]卸载插件 
echo   [3]查看状态 
echo   [4]退出 
echo.
echo ========================================

set /p choice=请选择操作 (1-4): 

if "%choice%"=="1" goto install
if "%choice%"=="2" goto uninstall
if "%choice%"=="3" goto status
if "%choice%"=="4" exit /b
goto menu

:install
cls
echo ========================================
echo        安装 PowerPoint 插件 
echo ========================================
setlocal enabledelayedexpansion
set existnet=false 
for /f "tokens=7 delims=\" %%a in ('REG QUERY HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319\SKUs') do (
	if not "%%a" == "Client" if not "%%a" == "Default" (
		if "%1" == "" (
			SET gg=%%a			
			SET ss=!gg:~22,6!
			SET ss=!ss:,=!
			SET ss=!ss:P=!
			rem 打印.net framework版本号
			rem echo !ss!
			if "!ss:v4.8=!" NEQ "!ss!" (
				set existnet=true
			)
		) else (			
			if "!%%a:v4.8=!" NEQ "!%%a!" (
			 	set existnet=true
			)	
			goto exit			
		)
	)
)
:exit
if %existnet% == true (
	echo   √ 运行环境检测完毕,已安装 .net framework v4.8运行环境
) else (
	echo 缺少运行环境按任意键退出,安装环境后继续执行安装...
	start NDP48.exe
	pause>nul
exit
 )
:: [1/6] 检查VSTO文件...
if exist "%~dp0PowerPointAddIn.vsto" (
    echo   √ 找到 PowerPointAddIn.vsto 
) else (
    echo   ❌ 错误: 未找到 PowerPointAddIn.vsto
    echo   请确保该文件与脚本在同一目录下。 
    echo 按下任意键返回菜单... &pause>nul
    goto menu
)

:: [2/6] 注册到HKCU...
set "REG_PATH=HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn"
set "MANIFEST=file:///%~dp0PowerPointAddIn.vsto|vstolocal"

reg add "%REG_PATH%" /f >nul 2>&1 && echo   √ HKCU-PPT  || echo   Χ HKCU-PPT创建失败 

:: [3/6] 设置插件参数...
reg add "%REG_PATH%" /v "FriendlyName" /t REG_SZ /d "网页插件" /f >nul 2>&1 && echo   √  -成功添加HKCU-FriendlyName  || echo   Χ HKCU-FriendlyName创建失败 
reg add "%REG_PATH%" /v "Description" /t REG_SZ /d "一款全能的Office网页插件" /f >nul 2>&1 && echo   √  -成功添加HKCU-Description  || echo   Χ HKCU-Description创建失败
reg add "%REG_PATH%" /v "LoadBehavior" /t REG_DWORD /d "3" /f >nul 2>&1 && echo   √  -成功添加HKCU-LoadBehavior  || echo   Χ HKCU-LoadBehavior创建失败
reg add "%REG_PATH%" /v "CommandLineSafe" /t REG_DWORD /d "1" /f >nul 2>&1 && echo   √  -成功添加HKCU-CommandLineSafe  || echo   Χ HKCU-CommandLineSafe创建失败
reg add "%REG_PATH%" /v "Manifest" /t REG_SZ /d "%MANIFEST:\=\\%" /f >nul 2>&1 && echo   √  -成功添加HKCU-Manifest  || echo   Χ HKCU-Manifest创建失败

:: [4/6] 注册到HKLM...
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

reg import "%TEMP%\ppt_system.reg" >nul 2>&1 && echo   √ HKLM-PPT-64位兼容  || echo   ⚠ HKLM-PPT-32位兼容安装失败

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

reg import "%TEMP%\ppt_wow64.reg" >nul 2>&1 && echo   √ HKLM-PPT-32位兼容  || echo   ⚠ HKLM-PPT-32位兼容失败 

:: [5/6] 注册到WPS Office...
reg add "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL" /v "PowerPointAddIn" /t REG_SZ /d "" /f >nul 2>&1
if errorlevel 1 (
    echo   ⚠ WPS注册失败
) else (
    echo   √ WPS Office兼容
)

:: [6/6] 清理临时文件...
del "%TEMP%\ppt_system.reg" >nul 2>&1
del "%TEMP%\ppt_wow64.reg" >nul 2>&1

echo.
echo ========================================
echo ✅ 插件安装完成！
echo ========================================
echo 按下任意键返回菜单... &pause>nul
goto menu

:uninstall
cls
echo ========================================
echo        卸载 PowerPoint 插件 
echo ========================================
echo.
echo 警告：此操作将删除插件注册信息。 
echo.
set /p confirm=确认卸载？(Y/N): 

if /i "%confirm%" neq "Y" goto menu

:: 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' neq '0' (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs -ArgumentList 'uninstall'"
    exit /b
)

echo.
echo 正在卸载插件...
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f>nul 2>&1 && echo   √HKCU-PPT删除成功    || echo   Χ HKCU-PPT删除失败！可能是不存在？ 
reg delete "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL" /f>nul 2>&1 && echo   √HKCU-WPS 兼容删除成功    || echo   Χ HKCU-WPS 兼容删除失败！可能是不存在？ 
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f>nul 2>&1 && echo   √HKLM-PPT-64删除成功    || echo   Χ HKLM-PPT-64删除失败！可能是不存在？ 
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f>nul 2>&1 && echo   √HKLM-PPT-32删除成功    || echo   Χ HKLM-PPT-32删除失败！可能是不存在？ 
"C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Uninstall file:///%~dp0PowerPointAddIn.vsto /s >nul 2>&1 && echo   √APP-PPT删除成功    || echo   Χ APP-PPT删除失败！可能是没使用双击vsto文件的方法安装过 


echo 按下任意键返回菜单... &pause>nul
goto menu

:status
cls
echo ========================================
echo        插件状态检查 
echo ========================================
reg query "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn">nul 2>&1 && echo   √HKCU-PPT已安装成功    || echo   Χ HKCU-PP不存在 
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn">nul 2>&1 && echo   √HKLM-PPT-64位已安装成功    || echo   Χ HKCU-PPT-64位不存在 
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn">nul 2>&1 && echo   √HKLM-PPT-32位已安装成功    || echo   Χ HKCU-PPT-32位不存在 
reg query "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL">nul 2>&1 && echo   √HKCU-WPS已安装成功    || echo   Χ HKCU-WPS不存在 
echo ========================================
echo 按下任意键返回菜单... &pause>nul
goto menu
pause