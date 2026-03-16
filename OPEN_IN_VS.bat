@echo off
REM MailMergeKit - Quick Start Script
REM This script opens the solution in Visual Studio 2022

echo.
echo ================================================
echo   MailMergeKit v0.0.1 - Development Launcher
echo ================================================
echo.

REM Try to find Visual Studio 2022
set VS2022="C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"
set VS2022PRO="C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\devenv.exe"
set VS2022ENT="C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"

if exist %VS2022% (
    echo Opening in Visual Studio 2022 Community...
    start "" %VS2022% "%~dp0MailMergeKit.sln"
    goto :END
)

if exist %VS2022PRO% (
    echo Opening in Visual Studio 2022 Professional...
    start "" %VS2022PRO% "%~dp0MailMergeKit.sln"
    goto :END
)

if exist %VS2022ENT% (
    echo Opening in Visual Studio 2022 Enterprise...
    start "" %VS2022ENT% "%~dp0MailMergeKit.sln"
    goto :END
)

REM If Visual Studio not found, open with default handler
echo Visual Studio 2022 not found in default locations.
echo Opening solution with default program...
echo.
start "" "%~dp0MailMergeKit.sln"

:END
echo.
echo ================================================
echo   Next Steps:
echo ================================================
echo   1. Press F5 to build and run
echo   2. Word will launch with MailMergeKit loaded
echo   3. Go to Mailings tab
echo   4. Look for "Send via MailMergeKit" button
echo.
echo   See BUILD_SUMMARY.md for detailed instructions
echo ================================================
echo.
pause
