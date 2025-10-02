@echo off
setlocal enabledelayedexpansion

:: ===== Guard: dotnet in PATH? =====
where dotnet >nul 2>nul
if errorlevel 1 (
  echo [ERROR] The 'dotnet' command was not found in PATH.
  echo         Install .NET SDK 8+ and/or add "C:\Program Files\dotnet" to PATH.
  echo         Then open a NEW cmd window and re-run this script.
  pause
  exit /b 1
)

:: ===== Ask for base path =====
set /p BASE_PATH=Enter the full path where you want to create a new directory (default C:\Users\sergeyr\WEB_QA_TESTS): 
if "%BASE_PATH%"=="" set "BASE_PATH=C:\Users\sergeyr\WEB_QA_TESTS"

:: Trim trailing backslash
if "%BASE_PATH:~-1%"=="\" set "BASE_PATH=%BASE_PATH:~0,-1%"

if not exist "%BASE_PATH%" (
  echo Path "%BASE_PATH%" doesn't exist. Creating it...
  mkdir "%BASE_PATH%" || ( echo [ERROR] Failed to create base path.& exit /b 1 )
)

:: ===== Ask for project folder name =====
set /p PROJECT_NAME=Enter the name of the new directory (project folder): 
if "%PROJECT_NAME%"=="" ( echo [ERROR] Directory name cannot be empty. & exit /b 1 )

:: ===== Create and cd =====
set "FULL_PROJECT_PATH=%BASE_PATH%\%PROJECT_NAME%"
if not exist "%FULL_PROJECT_PATH%" (
  mkdir "%FULL_PROJECT_PATH%" || ( echo [ERROR] Failed to create "%FULL_PROJECT_PATH%".& exit /b 1 )
)
cd /d "%FULL_PROJECT_PATH%" || ( echo [ERROR] Cannot cd to "%FULL_PROJECT_PATH%".& exit /b 1 )
echo Current directory: %FULL_PROJECT_PATH%

:: ===== Create NUnit project at root =====
dotnet new nunit --force
if errorlevel 1 ( echo [ERROR] dotnet new failed.& exit /b 1 )

:: Find the csproj file name
set "CSPROJ="
for %%F in (*.csproj) do set "CSPROJ=%%~nxF"
if "%CSPROJ%"=="" (
  echo [ERROR] No .csproj file found after template creation.
  exit /b 1
)
echo Project file: %CSPROJ%

:: ===== Upgrade NUnit then add NUnitLite to avoid NU1605 =====
echo Upgrading NUnit to 4.4.0...
dotnet add "%CSPROJ%" package NUnit --version 4.4.0 || ( echo [ERROR] Failed to add/upgrade NUnit.& exit /b 1 )
echo Adding NUnitLite 4.4.0...
dotnet add "%CSPROJ%" package NUnitLite --version 4.4.0 || ( echo [ERROR] Failed to add NUnitLite.& exit /b 1 )

:: ===== Add other packages =====
echo Adding other packages...
dotnet add "%CSPROJ%" package Dapper
dotnet add "%CSPROJ%" package System.Data.SqlClient
dotnet add "%CSPROJ%" package Twilio
dotnet add "%CSPROJ%" package Google.Apis.Gmail.v1

:: ===== Make project an EXE (OutputType) — single safe PowerShell replace =====
echo Converting project to EXE OutputType...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p = Get-Content '%CSPROJ%' -Raw;" ^
  "if ($p -notmatch '<StartupObject>') {" ^
  "  $p = $p -replace '</Project>', '<PropertyGroup><StartupObject>TestProcedure</StartupObject></PropertyGroup></Project>';" ^
  "} else {" ^
  "  $p = $p -replace '<StartupObject>.*?</StartupObject>', '<StartupObject>TestProcedure</StartupObject>';" ^
  "}" ^
  "Set-Content '%CSPROJ%' -Value $p -Encoding UTF8"

if errorlevel 1 (
  echo [ERROR] Failed to modify %CSPROJ%
  exit /b 1
)

:: ===== Create TestProcedure.cs (NUnitLite runner) — one-line PowerShell (no ^) =====
echo Creating TestProcedure.cs with NUnitLite runner...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^ "$lines = @('using NUnitLite;','','class TestProcedure','{','    public static int Main(string[] args)','    {','        return new AutoRun().Execute(args);','    }','}'); Set-Content -Path 'TestProcedure.cs' -Value $lines -Encoding UTF8"

if errorlevel 1 ( echo [ERROR] Failed to create TestProcedure.cs & exit /b 1 )

:: ===== Create Tests.cs (sample test) — one-line PowerShell (no ^) =====
echo Creating Tests.cs sample test...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^ "$lines = @('using NUnit.Framework;','','public class Tests','{','    [Test]','    public void Sanity()','    {','    Assert.Pass(\"Sanity OK\");','    }','}'); Set-Content -Path 'Tests.cs' -Value $lines -Encoding UTF8"

if errorlevel 1 ( echo [ERROR] Failed to create Tests.cs & exit /b 1 )

echo Files created:
dir /b *.cs

:: ===== Build =====
echo Restoring and building...
dotnet build
if errorlevel 1 ( echo [ERROR] Build failed.& exit /b 1 )

echo.
echo Setup complete.
echo Project location: %FULL_PROJECT_PATH%
:: ===== Go to project directory =====
cd %FULL_PROJECT_PATH%
:: ===== Open the project =====
dotnet run
echo.
echo To list tests: dotnet run -- --explore
echo To run all: dotnet run
echo To filter: dotnet run -- --where "Name~Sanity"
echo.

pause
endlocal
