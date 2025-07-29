@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

cd /d "%~dp0"

chcp 936

echo:
echo =============================
echo 欢迎使用网页插件
echo =============================
echo 正在检测插件运行环境...
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
	echo 运行环境检测完毕,已安装 .net framework v4.8运行环境
) else (
	echo 缺少运行环境按任意键退出,安装环境后继续执行安装...
	start https://dotnet.microsoft.com/zh-cn/download/dotnet-framework/thank-you/net48-offline-installer
	pause>nul
exit
 )
echo 正在添加Office注册表...
rem 注册表操作: /v 子项名称 /t 数据类型 /d 数据 /f 不提示强行修改
echo 注册PPT插件...
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "FriendlyName" /t REG_SZ /d "网页插件" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "Description" /t REG_SZ /d "一款全能的Office网页插件" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "LoadBehavior" /t REG_DWORD /d "3" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "CommandLineSafe" /t REG_DWORD /d "1" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "Manifest" /t REG_SZ /d "file:///%~dp0PowerPointAddIn.vsto|vstolocal" /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /s /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /s /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL" /v "PowerPointAddIn" /t REG_SZ /d "" /f
echo:
echo 插件安装完毕!!
pause