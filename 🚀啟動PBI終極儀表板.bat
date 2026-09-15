@echo off
title PBI AI Bridge

rem ---------------------------------------------------------------------------
rem  NOTE: keep this file ASCII-only.
rem  cmd.exe tracks its position in the batch file in BYTES. Mixing "chcp 65001"
rem  with multi-byte characters makes it lose track and start executing the
rem  middle of a line as a command. English messages are the safe choice here.
rem ---------------------------------------------------------------------------

set "ROOT=%~dp0"
set "PROJECT=%ROOT%pbibridge_csharp"
set "DLL=%PROJECT%\bin\Release\net8.0\pbibridge_csharp.dll"
set "CONFIG=%PROJECT%\appsettings.json"
set "TEMPLATE=%PROJECT%\appsettings.template.json"
set "HTML=%ROOT%PowerBI_Visualizer.html"

rem Prefer the dotnet under the user profile, fall back to PATH
if exist "%USERPROFILE%\.dotnet\dotnet.exe" (
    set "DOTNET=%USERPROFILE%\.dotnet\dotnet.exe"
) else (
    set "DOTNET=dotnet"
)

echo ==============================================
echo    PBI AI Bridge
echo ==============================================
echo.

rem --- [1/4] config ---------------------------------------------------------
rem First run has no appsettings.json - build one from the template and put a
rem freshly generated API Key in it. Every machine gets its own key; do not share.
if not exist "%CONFIG%" goto MAKE_CONFIG
echo [1/4] Config file      OK
goto CONFIG_DONE

:MAKE_CONFIG
if not exist "%TEMPLATE%" goto NO_TEMPLATE
echo [1/4] Config file      missing - creating it for you...
echo.
rem Paths go through the environment (Unicode) instead of being pasted into the
rem PowerShell command line, so a non-ASCII folder name cannot mangle them.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$k=[guid]::NewGuid().ToString('N').ToUpper(); $t=[IO.File]::ReadAllText($env:TEMPLATE); [IO.File]::WriteAllText($env:CONFIG, $t.Replace('__PUT_YOUR_OWN_RANDOM_KEY_HERE__',$k), (New-Object Text.UTF8Encoding $false)); Write-Host ('      Your API Key: ' + $k)"
if errorlevel 1 goto CONFIG_FAIL
if not exist "%CONFIG%" goto CONFIG_FAIL
echo.
echo       Created: %CONFIG%
echo       The web UI will ask for that key once, then remember it.
echo       You can look it up again in appsettings.json at any time.
echo.
pause

:CONFIG_DONE

rem --- [2/4] port 5500 ------------------------------------------------------
rem A previous instance left running is the usual cause of "address already in use"
netstat -ano | findstr /r /c:":5500 .*LISTENING" >nul 2>&1
if %errorlevel% equ 0 goto PORT_BUSY
echo [2/4] Port 5500        free

rem --- [3/4] build ----------------------------------------------------------
echo [3/4] Building (Release)...
"%DOTNET%" build "%PROJECT%" -c Release --nologo -v q
if errorlevel 1 goto BUILD_FAIL
if not exist "%DLL%" goto BUILD_FAIL
echo       Build OK

rem --- [4/4] launch ---------------------------------------------------------
echo [4/4] Opening web UI and starting server...
start "" "%HTML%"
echo.
echo ----------------------------------------------
echo   Server is running - KEEP THIS WINDOW OPEN!
echo   Closing this window stops the server.
echo ----------------------------------------------
echo.

"%DOTNET%" "%DLL%"

echo.
echo [EXIT] Server stopped.
pause
exit /b 0


:NO_TEMPLATE
echo [ERROR] appsettings.template.json not found.
echo         Expected at: %TEMPLATE%
echo         The copy you received is incomplete - ask for the file.
echo.
pause
exit /b 1


:CONFIG_FAIL
echo.
echo [ERROR] Could not create appsettings.json automatically.
echo         Do it by hand: copy appsettings.template.json to appsettings.json,
echo         then replace __PUT_YOUR_OWN_RANDOM_KEY_HERE__ with the output of
echo           [System.Guid]::NewGuid().ToString('N').ToUpper()
echo.
pause
exit /b 1


:PORT_BUSY
echo.
echo [ERROR] Port 5500 is already in use - a server is probably already running.
echo.
echo         List the owners (PowerShell):
echo           Get-NetTCPConnection -LocalPort 5500 -State Listen
echo.
echo         Note: there is sometimes more than one instance. All of them must
echo         be stopped before the port is released.
echo.
pause
exit /b 1


:BUILD_FAIL
echo.
echo [ERROR] Build failed - server not started.
echo         Scroll up to see the compiler errors.
echo.
pause
exit /b 1
