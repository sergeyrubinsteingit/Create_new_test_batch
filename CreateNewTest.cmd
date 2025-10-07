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

:: ===== Find the csproj file name =====
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
dotnet add "%CSPROJ%" package OpenQA.Selenium
dotnet add "%CSPROJ%" package Dapper
dotnet add "%CSPROJ%" package System.Data.SqlClient
dotnet add "%CSPROJ%" package Twilio
dotnet add "%CSPROJ%" package Google.Apis.Gmail.v1
dotnet add "%CSPROJ%" package Microsoft.Office.Interop.Outlook

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

:: ===== Create TestProcedure.cs — one-line PowerShell (no ^) =====
echo Creating TestProcedure.cs...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^ "$lines = @(   'using System;','using System.Net.Http;','using System.Net.Http.Headers;','using System.Text;','using System.Text.Json;','using System.Threading;','using System.Threading.Tasks;','','','public class TestProcedure','{','    private readonly HttpClient _http;','    private readonly JsonSerializerOptions _json = new(JsonSerializerDefaults.Web);','    private string? _token;','','    // Normal instance constructor with parameter','    public TestProcedure(HttpClient? httpClient = null)','    {','        _http = httpClient ?? new HttpClient { BaseAddress = new Uri(\"https://qa-cortex.nayax.com/\") };','    }','','    public static async Task Main(string[] args)','    {','        var _procedure = new TestProcedure();   // Instantiate','        await _procedure.RunAsync();','    }','','    public async Task RunAsync()','    {','        // Token check-in','        try','        {','            string token = await SignIn(\"sergeyr\", \"rubi69qa1******\");','            Console.WriteLine($\"Received token in Sign in: {token}\");','        }','        catch (Exception ex)','        {','            throw new Exception($\"Sign-in failed, test is aborted:\n{ex.Message}\");','        }','','        // Test 1:','        try','        {','            if (string.IsNullOrEmpty(_token))','            {','				throw new Exception(\"Token is null, test is aborted.\"); ','            }','            else ','			{','                string token = await NextTestStep(_token);','            }','        }','        catch (Exception ex)','        {','            Console.WriteLine($\"Function failed: {ex.Message}\");','        }','','        _http.Dispose();','    }','','    // Sign-in at start of the flow','    public async Task<string> SignIn(string username, string password, CancellationToken cancelToken = default)','    {','        var url = \"users/v1/signin\";','        var payload = new { username, password };','        var body = new StringContent(JsonSerializer.Serialize(payload, _json), Encoding.UTF8, \"application/json\");','','        using var rsp = await _http.PostAsync(url, body, cancelToken).ConfigureAwait(false);','        rsp.EnsureSuccessStatusCode();','','        var json = await rsp.Content.ReadAsStringAsync(cancelToken).ConfigureAwait(false);','        using var doc = JsonDocument.Parse(json);','        var tokenProp = doc.RootElement.GetProperty(\"token\");','','        _token = tokenProp.GetString();','','        if (string.IsNullOrEmpty(_token))','            throw new InvalidOperationException(\"Token not found in sign-in response.\");','','        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(\"Bearer\", _token);','        return _token;','    }','','    // Test 1','    public async Task<string> NextTestStep(string token)','    {','        // If the token OK?','        await EnsureSignedIn(token);','        Console.WriteLine($\"Received token in the Next Test Step:\n{token}\");','        return token;','    }','','    // Sign-in check in before running the next step','    private async Task<string> EnsureSignedIn(string controlToken)','    {','        _token = controlToken;','        if (string.IsNullOrWhiteSpace(_token))','            throw new InvalidOperationException(\"Not signed in. Call SignIn first.\");','        return _token;','    }','}'   ); Set-Content -Path 'TestProcedure.cs' -Value $lines -Encoding UTF8"

if errorlevel 1 ( echo [ERROR] Failed to create TestProcedure.cs & exit /b 1 )

:: ===== Create Tests.cs (sample test) — one-line PowerShell (no ^) =====
echo Creating Tests.cs sample test...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^ "$lines = @('using NUnit.Framework;','using System.Net.Http;', 'using System.Net.Http.Headers;', 'using System.Text;', 'using System.Text.Json;','', 'public sealed class LoginRequest', '{', '    public string Username { get; set; } = \"\";', '    public string Password { get; set; } = \"\";', '}', '','public class Tests','{', '    [Test]','    public void SignIn()','    {','    Assert.Pass(\"Sanity OK\");','    }','}'); Set-Content -Path 'Tests.cs' -Value $lines -Encoding UTF8"

if errorlevel 1 ( echo [ERROR] Failed to create Tests.cs & exit /b 1 )

echo Files created:
dir /b *.cs

:: ===== Build =====
echo Restoring and building...
dotnet build
if errorlevel 1 ( echo [ERROR] Build failed.& exit /b 1 )

echo.
echo Setup complete.

:: ===== Go to project directory =====
echo Project location: %FULL_PROJECT_PATH%
cd %FULL_PROJECT_PATH%

:: ===== Open the project =====
dotnet run --project "%FULL_PROJECT_PATH%\%PROJECT_NAME%.csproj" %*
echo.
echo To list tests: dotnet run -- --explore
echo To run all: dotnet run
echo To filter: dotnet run -- --where "Name~Sanity"
echo.

:: ===== Ask to open in Visual Studio =====
set /p OPEN_VS=Open the project in Visual Studio? [y/n]: 
if "%OPEN_VS%"=="y" (
    echo Opening project: %CSPROJ%
    start "" "%CSPROJ%"
) else (
    echo Project creation complete. Visual Studio not opened.
)

pause
endlocal
