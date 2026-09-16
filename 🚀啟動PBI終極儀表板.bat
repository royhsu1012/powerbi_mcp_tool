@echo off
setlocal
title PBI AI Bridge

rem ---------------------------------------------------------------------------
rem  PBI AI Bridge - the only file you ever double-click.
rem
rem  Every run walks the same checklist and only does what is missing:
rem    already running?  -> just open the dashboard
rem    .NET SDK          -> first time only (offers to install it)
rem    settings file     -> first time only (creates your own API Key)
rem    libraries         -> first time only (offers to download them)
rem    compile           -> only when the program changed
rem    start             -> the server opens the dashboard once it is ready
rem
rem  NOTE: keep this file ASCII-only with CRLF line endings.
rem  cmd.exe tracks its position in a batch file in BYTES. Multi-byte text or
rem  LF-only line endings make it lose its place, and goto / labels break.
rem ---------------------------------------------------------------------------

set "ROOT=%~dp0"
set "PROJECT=%ROOT%pbibridge_csharp"
set "CONFIG=%PROJECT%\appsettings.json"
set "TEMPLATE=%PROJECT%\appsettings.template.json"
set "ASSETS=%PROJECT%\obj\project.assets.json"
set "BIN=%PROJECT%\bin\Release\net8.0"
set "DLL=%BIN%\pbibridge_csharp.dll"
set "STAMP=%BIN%\build.stamp"
set "URL=http://localhost:5500/"
set "SDKLIST=%TEMP%\pbi_bridge_sdks.txt"

echo ==============================================
echo    PBI AI Bridge
echo ==============================================
echo.

rem === already running? =======================================================
rem Double-clicking again while the server runs is common. Answer with the
rem dashboard instead of a compile error about locked files.
netstat -ano | findstr /r /c:":5500 .*LISTENING" >nul 2>&1
if errorlevel 1 goto STEP_FILES
powershell -NoProfile -Command "try { if ((Invoke-RestMethod -Uri 'http://localhost:5500/ping' -TimeoutSec 3) -eq 'pong') { exit 0 } } catch { }; exit 1" >nul 2>&1
if errorlevel 1 goto PORT_BUSY
echo  The server is already running in another window.
echo  Opening the dashboard - this window closes by itself.
start "" "%URL%"
timeout /t 5 >nul
exit /b 0


rem === [1/5] files ============================================================
:STEP_FILES
echo [1/5] Files
if not exist "%PROJECT%\Program.cs" goto INCOMPLETE
if not exist "%TEMPLATE%" goto INCOMPLETE
if not exist "%CONFIG%" call :LOCATION_CHECK
echo       OK


rem === [2/5] .NET SDK =========================================================
echo [2/5] .NET SDK
call :FIND_DOTNET
if not defined DOTNET goto NO_SDK
:SDK_OK
echo       OK


rem === [3/5] settings file ====================================================
echo [3/5] Settings file
if exist "%CONFIG%" goto CONFIG_OK
rem Paths travel through the environment (Unicode) instead of the command line,
rem so a folder name with non-ASCII characters cannot mangle them.
powershell -NoProfile -Command "$k=[guid]::NewGuid().ToString('N').ToUpper(); $t=[IO.File]::ReadAllText($env:TEMPLATE); [IO.File]::WriteAllText($env:CONFIG, $t.Replace('__PUT_YOUR_OWN_RANDOM_KEY_HERE__',$k), (New-Object Text.UTF8Encoding $false))"
if errorlevel 1 goto CONFIG_FAIL
if not exist "%CONFIG%" goto CONFIG_FAIL
echo       Created, with an API Key made just for this computer.
echo       The dashboard and the AI tools read it by themselves.
call :PROTECTION_REMINDER
goto STEP_LIBS
:CONFIG_OK
echo       OK


rem === [4/5] libraries ========================================================
:STEP_LIBS
echo [4/5] Libraries
if exist "%ASSETS%" goto LIBS_OK
echo       First run: 6 libraries must be downloaded from nuget.org, about 17 MB.
echo       They are Microsoft's Analysis Services client libraries. Only this once.
echo.
choice /c YN /n /m "      Download them now? [Y/N] "
if errorlevel 2 goto LIBS_DECLINED
:LIBS_RETRY
echo       Downloading...
"%DOTNET%" restore "%PROJECT%" --nologo -v q
if errorlevel 1 goto RESTORE_FAIL
:LIBS_OK
echo       OK


rem === [5/5] compile and start ================================================
echo [5/5] Compile
if not exist "%DLL%" goto DO_BUILD
if not exist "%STAMP%" goto DO_BUILD
rem Compile only when the program changed: faster start, and no freshly written
rem DLL for the antivirus to inspect on every launch.
powershell -NoProfile -Command "$s=(Get-Item -LiteralPath $env:STAMP).LastWriteTimeUtc; foreach ($f in 'Program.cs','pbibridge_csharp.csproj') { if ((Get-Item -LiteralPath (Join-Path $env:PROJECT $f)).LastWriteTimeUtc -gt $s) { exit 1 } }; exit 0"
if errorlevel 1 goto DO_BUILD
echo       Up to date
goto START

:DO_BUILD
echo       Compiling, about 30 seconds...
"%DOTNET%" build "%PROJECT%" -c Release --nologo -v q
if errorlevel 1 goto BUILD_FAIL
if not exist "%DLL%" goto BUILD_FAIL
type nul > "%STAMP%"
echo       OK

:START
rem Always refresh the settings copy the server reads. A plain copy is cheap and
rem also covers a settings file restored from an older backup, which the
rem compiler's newer-than check would skip.
copy /y "%CONFIG%" "%BIN%\appsettings.json" >nul
echo.
echo ----------------------------------------------
echo   Starting - the dashboard opens by itself.
echo   KEEP THIS WINDOW OPEN. Closing it stops the server.
echo   Dashboard address: %URL%
echo ----------------------------------------------
echo.
"%DOTNET%" "%DLL%"
echo.
echo [EXIT] Server stopped. If that was unexpected, scroll up for the reason.
pause
exit /b 0


rem ===========================================================================
rem  Problems - each one explains what happened and what to do
rem ===========================================================================

:INCOMPLETE
echo.
echo [ERROR] Some files are missing from this folder.
echo         Expected: pbibridge_csharp\Program.cs and appsettings.template.json
echo.
echo         The copy you received is incomplete. Get the whole folder again,
echo         and unzip it first - do not run it from inside the zip.
echo.
pause
exit /b 1


:NO_SDK
echo.
echo       Not found. This tool needs the .NET SDK, version 8 or newer.
echo       The .NET Runtime alone is not enough - it cannot compile.
echo.
where winget >nul 2>&1
if errorlevel 1 goto SDK_MANUAL
echo       It can be installed right here with winget, the installer built into Windows:
echo         Package   Microsoft .NET SDK 8   id: Microsoft.DotNet.SDK.8
echo         Source    Microsoft, about 210 MB
echo         Note      Windows may ask for administrator permission.
echo.
choice /c YN /n /m "      Install it now? [Y/N] "
if errorlevel 2 goto SDK_MANUAL
echo.
winget install --id Microsoft.DotNet.SDK.8 --exact --source winget --accept-package-agreements --accept-source-agreements
call :FIND_DOTNET
if defined DOTNET goto SDK_OK
echo.
echo       [ERROR] The install did not finish, or it was cancelled.

:SDK_MANUAL
echo.
echo       To install it by hand:
echo         1. Open  https://dotnet.microsoft.com/download/dotnet/8.0
echo         2. In the SDK 8.0.x column, pick  Windows  x64  Installer
echo         3. Run the installer, then come back to this window and press R
echo.
:SDK_WAIT
choice /c ORQ /n /m "      O = open that page,  R = check again,  Q = quit : "
if errorlevel 3 exit /b 1
if errorlevel 2 goto SDK_RECHECK
start "" "https://dotnet.microsoft.com/download/dotnet/8.0"
goto SDK_WAIT
:SDK_RECHECK
call :FIND_DOTNET
if defined DOTNET goto SDK_OK
echo       Still not found. If the installer is still running, wait for it to
echo       finish and press R again.
goto SDK_WAIT


:CONFIG_FAIL
echo.
echo [ERROR] Could not create pbibridge_csharp\appsettings.json.
echo.
echo         Do it by hand: copy appsettings.template.json to appsettings.json,
echo         then replace __PUT_YOUR_OWN_RANDOM_KEY_HERE__ with the output of
echo           [System.Guid]::NewGuid().ToString('N').ToUpper()
echo.
pause
exit /b 1


:LIBS_DECLINED
echo.
echo       The server cannot be compiled without these libraries.
echo       Double-click this file again whenever you are ready.
echo.
pause
exit /b 1


:RESTORE_FAIL
echo.
echo       [ERROR] The download failed.
"%DOTNET%" nuget list source | findstr /i /c:"api.nuget.org" >nul
if not errorlevel 1 goto RESTORE_NETWORK
echo.
echo       Cause: nuget.org is not registered as a package source on this computer.
choice /c YN /n /m "      Add nuget.org as a package source and try again? [Y/N] "
if errorlevel 2 goto RESTORE_NETWORK
"%DOTNET%" nuget add source https://api.nuget.org/v3/index.json -n nuget.org
goto LIBS_RETRY

:RESTORE_NETWORK
echo.
echo       Most likely the network cannot reach nuget.org. Check:
echo         1. In a browser, open  https://api.nuget.org/v3/index.json
echo            You should see a page of text.
echo         2. The browser works but this still fails: usually a company proxy.
echo            Ask IT for help, or for an internal NuGet package source.
echo         3. On a VPN? Try again with it switched the other way.
echo.
choice /c RQ /n /m "      R = try again,  Q = quit : "
if errorlevel 2 exit /b 1
goto LIBS_RETRY


:BUILD_FAIL
echo.
echo [ERROR] Compile failed - scroll up for the compiler errors.
echo.
echo         being used by another process
echo             Another launcher window is compiling right now.
echo             Close the extra window, then double-click this file again.
echo         NU1301
echo             The libraries could not be downloaded - check the network.
echo.
pause
exit /b 1


:PORT_BUSY
echo.
echo [ERROR] Port 5500 is used by a different program.
echo         VS Code Live Server uses 5500 by default, for example.
echo         Close that program, then double-click this file again.
echo.
pause
exit /b 1


rem ===========================================================================
rem  Subroutines
rem ===========================================================================

rem Sets DOTNET to the first dotnet that has an SDK of version 8 or newer.
rem Checks the fixed install folders too: right after a fresh install, PATH in
rem this window is still the old one.
:FIND_DOTNET
set "DOTNET="
if exist "%USERPROFILE%\.dotnet\dotnet.exe" call :TRY_DOTNET "%USERPROFILE%\.dotnet\dotnet.exe"
if not defined DOTNET if exist "%ProgramFiles%\dotnet\dotnet.exe" call :TRY_DOTNET "%ProgramFiles%\dotnet\dotnet.exe"
if not defined DOTNET call :TRY_DOTNET dotnet
exit /b 0

:TRY_DOTNET
set "SDK_OK_FLAG="
"%~1" --list-sdks > "%SDKLIST%" 2>nul
for /f "usebackq tokens=1 delims=." %%V in ("%SDKLIST%") do if %%V GEQ 8 set "SDK_OK_FLAG=1"
del "%SDKLIST%" >nul 2>&1
if defined SDK_OK_FLAG set "DOTNET=%~1"
exit /b 0


:LOCATION_CHECK
echo "%ROOT%" | findstr /i "OneDrive" >nul
if errorlevel 1 exit /b 0
echo.
echo       [WARNING] This folder is inside OneDrive. Sync can lock files while
echo       compiling and make it fail at random. A local folder such as
echo       C:\PBI_AI_Bridge is strongly recommended.
echo.
pause
exit /b 0


:PROTECTION_REMINDER
echo.
echo       ------------------------------------------------------------------
echo       IMPORTANT - protect your data before AI queries it
echo.
echo       The default list only knows English column names such as
echo       customer, amount, price. If your model uses Chinese column names,
echo       add them to the DataProtection section of
echo         pbibridge_csharp\appsettings.json
echo       See README.md, section 4, for examples.  
echo       ------------------------------------------------------------------
echo.
choice /c YN /n /m "      Open that file in Notepad now? [Y/N] "
if errorlevel 2 exit /b 0
start "" notepad "%CONFIG%"
echo       Edit, save and close Notepad - then press any key here to continue.
pause >nul
exit /b 0
