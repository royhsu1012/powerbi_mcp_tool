@echo off
title PBI AI Bridge - Setup
setlocal enabledelayedexpansion

rem ---------------------------------------------------------------------------
rem  NOTE: keep this file ASCII-only.
rem  cmd.exe tracks its position in the batch file in BYTES. Non-ASCII text read
rem  under the OEM codepage turns into mojibake and can break paths. That is why
rem  the launcher is located by wildcard below instead of by its real name.
rem ---------------------------------------------------------------------------

set "ROOT=%~dp0"
set "PROJECT=%ROOT%pbibridge_csharp"
set "CONFIG=%PROJECT%\appsettings.json"
set "TEMPLATE=%PROJECT%\appsettings.template.json"
set "DLL=%PROJECT%\bin\Release\net8.0\pbibridge_csharp.dll"
set "SDKLIST=%TEMP%\pbi_bridge_sdks.txt"

rem The launcher's name contains an emoji, which cannot survive a round trip
rem through the OEM codepage - so we never put its path in a variable. We only
rem confirm it is there, by matching the ASCII "PBI" inside its name.
set "LAUNCHER_FOUND="
for %%F in ("%ROOT%*PBI*.bat") do set "LAUNCHER_FOUND=1"

echo.
echo ==================================================
echo    PBI AI Bridge - Setup
echo ==================================================
echo.
echo  This runs once. It checks your machine, creates
echo  your own API Key, and compiles the server.
echo.
echo  It never touches your Power BI files. The first
echo  compile downloads libraries from nuget.org once.
echo.
pause
echo.

rem === [1/5] project files ===================================================
echo [1/5] Checking the copy you received...
if not exist "%PROJECT%" goto INCOMPLETE
if not exist "%TEMPLATE%" goto INCOMPLETE
if not defined LAUNCHER_FOUND goto INCOMPLETE
echo       OK
echo.

rem === [2/5] .NET 8 SDK =======================================================
echo [2/5] Checking for the .NET SDK...
if exist "%USERPROFILE%\.dotnet\dotnet.exe" (
    set "DOTNET=%USERPROFILE%\.dotnet\dotnet.exe"
) else (
    set "DOTNET=dotnet"
)
"%DOTNET%" --version >nul 2>&1
if errorlevel 1 goto NO_DOTNET

rem "dotnet" alone can be Runtime-only. An empty SDK list means no SDK.
"%DOTNET%" --list-sdks > "%SDKLIST%" 2>nul
for %%A in ("%SDKLIST%") do if %%~zA equ 0 goto NO_SDK
echo       Found:
for /f "tokens=1" %%V in ('type "%SDKLIST%"') do echo         .NET SDK %%V
del "%SDKLIST%" >nul 2>&1
echo.

rem === [3/5] where this folder lives ==========================================
echo [3/5] Checking the folder location...
echo %ROOT% | findstr /i "OneDrive" >nul
if not errorlevel 1 goto ONEDRIVE_WARN
echo       OK
goto LOCATION_DONE

:ONEDRIVE_WARN
echo.
echo       [WARNING] This folder is inside OneDrive.
echo       Sync can lock files while the compiler writes them, which makes
echo       the build fail at random. Moving it somewhere local - for example
echo       C:\PBI_AI_Bridge - is strongly recommended.
echo.
echo       Setup will continue anyway.
echo.
pause

:LOCATION_DONE
echo.

rem === [4/5] config + API key =================================================
echo [4/5] Preparing your settings file...
if exist "%CONFIG%" goto CONFIG_EXISTS

rem Paths travel through the environment (Unicode) instead of the command line,
rem so a folder name with non-ASCII characters cannot mangle them.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$k=[guid]::NewGuid().ToString('N').ToUpper(); $t=[IO.File]::ReadAllText($env:TEMPLATE); [IO.File]::WriteAllText($env:CONFIG, $t.Replace('__PUT_YOUR_OWN_RANDOM_KEY_HERE__',$k), (New-Object Text.UTF8Encoding $false))"
if errorlevel 1 goto CONFIG_FAIL
if not exist "%CONFIG%" goto CONFIG_FAIL
echo       Created appsettings.json with a key generated just for you.
goto CONFIG_DONE

:CONFIG_EXISTS
echo       appsettings.json already exists - keeping it as is.

:CONFIG_DONE
echo.

rem === [5/5] build ============================================================
echo [5/5] Compiling the server (about 30 seconds, one time only)...
"%DOTNET%" build "%PROJECT%" -c Release --nologo -v q
if errorlevel 1 goto BUILD_FAIL
if not exist "%DLL%" goto BUILD_FAIL
echo       Build OK
echo.

rem === done ===================================================================
echo ==================================================
echo    Setup complete
echo ==================================================
echo.
echo  Your API Key (the web page asks for it once):
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$j=Get-Content -LiteralPath $env:CONFIG -Raw -Encoding UTF8 | ConvertFrom-Json; Write-Host ('      ' + $j.Security.ApiKey)"
echo.
echo  You can look it up again any time in:
echo      pbibridge_csharp\appsettings.json
echo.
echo  From now on, this is all you do:
echo      1. Open a Power BI file (PBIX or PBIP)
echo      2. Double-click the launcher (the file with the rocket icon)
echo      3. Keep that window open - closing it stops the server
echo.
echo  You will not need to run this setup again.
echo.
echo --------------------------------------------------
echo  Press any key - this folder will open so you can
echo  see the launcher. Double-click the rocket file.
echo --------------------------------------------------
pause >nul
start "" "%ROOT%"
exit /b 0


:INCOMPLETE
echo.
echo [ERROR] Some files are missing from this folder.
echo         Expected next to this script:
echo           pbibridge_csharp\  (with appsettings.template.json inside)
echo           the launcher .bat  (rocket icon)
echo.
echo         The copy you received is incomplete - ask for the whole folder
echo         again, and make sure you unzipped it rather than opening the zip.
echo.
pause
exit /b 1


:NO_DOTNET
echo.
echo [ERROR] The .NET SDK is not installed (or not on PATH).
echo.
echo         Opening the download page for you.
echo         Pick: .NET SDK 8.x.x  ^>  Windows  ^>  x64 Installer
echo         It must be the SDK. The Runtime alone cannot compile.
echo.
echo         Install it, then run this setup again.
echo.
start "" "https://dotnet.microsoft.com/download/dotnet/8.0"
pause
exit /b 1


:NO_SDK
del "%SDKLIST%" >nul 2>&1
echo.
echo [ERROR] Found the .NET Runtime but no SDK.
echo         This is the single most common setup problem.
echo.
echo         Opening the download page for you.
echo         Pick: .NET SDK 8.x.x  ^>  Windows  ^>  x64 Installer
echo.
echo         Install it, then run this setup again.
echo.
start "" "https://dotnet.microsoft.com/download/dotnet/8.0"
pause
exit /b 1


:CONFIG_FAIL
echo.
echo [ERROR] Could not create appsettings.json.
echo.
echo         Do it by hand: copy appsettings.template.json to appsettings.json,
echo         then replace __PUT_YOUR_OWN_RANDOM_KEY_HERE__ with the output of
echo           [System.Guid]::NewGuid().ToString('N').ToUpper()
echo.
pause
exit /b 1


:BUILD_FAIL
echo.
echo [ERROR] Compile failed - scroll up for the compiler errors.
echo.
echo         Most likely causes:
echo           - Runtime installed instead of SDK
echo           - no internet access to nuget.org (NU1301 above):
echo             check the network or company proxy, then run this again
echo           - folder sits in OneDrive and sync locked a file
echo           - the server is already running and holds the DLL
echo             (close its console window, then run this again)
echo.
pause
exit /b 1
