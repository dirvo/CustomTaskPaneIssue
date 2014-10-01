@echo off
setlocal
cd /d "%~dp0"

goto check_Permissions

:check_permissions
REM Quick test for Windows generation: UAC aware or not ; all OS before NT4 ignored for simplicity
SET NewOSWith_UAC=YES
VER | FINDSTR /IL "5." > NUL
IF %ERRORLEVEL% == 0 SET NewOSWith_UAC=NO
VER | FINDSTR /IL "4." > NUL
IF %ERRORLEVEL% == 0 SET NewOSWith_UAC=NO

CALL NET SESSION >nul 2>&1
IF NOT %ERRORLEVEL% == 0 (
	
	if /i "%NewOSWith_UAC%"=="YES" (
		echo Restarting script as administrator
        rem Start batch again with UAC
        echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
        echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
        
		cscript //nologo "%temp%\getadmin.vbs"
        del "%temp%\getadmin.vbs"
		exit /B
	)
	goto :eof
)

:register_addin

"%WINDIR%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase ComAddin\bin\Debug\ComAddin.dll
IF NOT %ERRORLEVEL% == 0 (goto error)

reg add HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\ComAddin.Connect /v Description /t REG_SZ /d "CustomTaskPaneIssue AddIn" /f
reg add HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\ComAddin.Connect /v FriendlyName /t REG_SZ /d "CustomTaskPaneIssue AddIn" /f
reg add HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\ComAddin.Connect /v LoadBehavior /t REG_DWORD /d 3 /f

goto :eof

:error
pause

