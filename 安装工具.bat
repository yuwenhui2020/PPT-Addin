@echo off

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

cd /d "%~dp0"

chcp 936

echo:
echo =============================
echo ��ӭʹ����ҳ���
echo =============================
echo ���ڼ�������л���...
setlocal enabledelayedexpansion
set existnet=false 
for /f "tokens=7 delims=\" %%a in ('REG QUERY HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319\SKUs') do (
	if not "%%a" == "Client" if not "%%a" == "Default" (
		if "%1" == "" (
			SET gg=%%a			
			SET ss=!gg:~22,6!
			SET ss=!ss:,=!
			SET ss=!ss:P=!
			rem ��ӡ.net framework�汾��
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
	echo ���л���������,�Ѱ�װ .net framework v4.8���л���
) else (
	echo ȱ�����л�����������˳�,��װ���������ִ�а�װ...
	start https://dotnet.microsoft.com/zh-cn/download/dotnet-framework/thank-you/net48-offline-installer
	pause>nul
exit
 )
echo �������Officeע���...
rem ע������: /v �������� /t �������� /d ���� /f ����ʾǿ���޸�
echo ע��PPT���...
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "FriendlyName" /t REG_SZ /d "��ҳ���" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "Description" /t REG_SZ /d "һ��ȫ�ܵ�Office��ҳ���" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "LoadBehavior" /t REG_DWORD /d "3" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "CommandLineSafe" /t REG_DWORD /d "1" /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /v "Manifest" /t REG_SZ /d "file:///%~dp0PowerPointAddIn.vsto|vstolocal" /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /s /f
reg copy "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn" /s /f
reg add "HKEY_CURRENT_USER\SOFTWARE\Kingsoft\Office\WPP\AddinsWL" /v "PowerPointAddIn" /t REG_SZ /d "" /f
echo:
echo �����װ���!!
pause