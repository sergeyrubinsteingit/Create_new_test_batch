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
dotnet add "%CSPROJ%" package OpenQA.Selenium.Chrome
dotnet add "%CSPROJ%" package Selenium.WebDriver
dotnet add "%CSPROJ%" package Selenium.Support
dotnet add "%CSPROJ%" package Dapper
rem dotnet add "%CSPROJ%" package System.Data.SqlClient
dotnet add "%CSPROJ%" package Microsoft.Data.SqlClient
dotnet add "%CSPROJ%" package Twilio
dotnet add "%CSPROJ%" package Google.Apis.Gmail.v1
dotnet add "%CSPROJ%" package Microsoft.Office.Interop.Outlook

:: ===== Make project an EXE (OutputType) — single safe PowerShell replace =====
echo Converting project to EXE OutputType...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p = Get-Content '%CSPROJ%' -Raw;" ^
  "if ($p -notmatch '<StartupObject>') {" ^
  "  $p = $p -replace '</Project>', '<PropertyGroup><StartupObject>MainClass</StartupObject></PropertyGroup></Project>';" ^
  "} else {" ^
  "  $p = $p -replace '<StartupObject>.*?</StartupObject>', '<StartupObject>MainClass</StartupObject>';" ^
  "}" ^
  "Set-Content '%CSPROJ%' -Value $p -Encoding UTF8"

if errorlevel 1 (
  echo [ERROR] Failed to modify %CSPROJ%
  exit /b 1
)

:: ===== Create MainClass.cs =====
echo Creating MainClass.cs sample test...

> "__mk_main.ps1" echo $content = @'
>> "__mk_main.ps1" echo using System.Text.Json;
>> "__mk_main.ps1" echo using HttpClient = System.Net.Http.HttpClient;
>> "__mk_main.ps1" echo //
>> "__mk_main.ps1" echo public class MainClass
>> "__mk_main.ps1" echo {
>> "__mk_main.ps1" echo     private readonly HttpClient _http;
>> "__mk_main.ps1" echo     private readonly JsonSerializerOptions _json = new(JsonSerializerDefaults.Web);
>> "__mk_main.ps1" echo     public string? _token;
>> "__mk_main.ps1" echo     private readonly SignIn _signIn;
>> "__mk_main.ps1" echo     private Test1 _test1;
>> "__mk_main.ps1" echo //
>> "__mk_main.ps1" echo     // Normal instance constructor with parameter
>> "__mk_main.ps1" echo     public MainClass(HttpClient? httpClient = null)
>> "__mk_main.ps1" echo     {
>> "__mk_main.ps1" echo         _http = httpClient ?? new HttpClient { BaseAddress = new Uri("https://qa-cortex.nayax.com/") };
>> "__mk_main.ps1" echo         _signIn = new SignIn(_http, _json);
>> "__mk_main.ps1" echo         _test1 = new Test1(_http);
>> "__mk_main.ps1" echo     }
>> "__mk_main.ps1" echo //
>> "__mk_main.ps1" echo     public static async Task Main(string[] args)
>> "__mk_main.ps1" echo     {
>> "__mk_main.ps1" echo         var _procedure = new MainClass();   // Instantiate
>> "__mk_main.ps1" echo         await _procedure.RunAsync();
>> "__mk_main.ps1" echo     }
>> "__mk_main.ps1" echo //
>> "__mk_main.ps1" echo     public async Task RunAsync()
>> "__mk_main.ps1" echo     {
>> "__mk_main.ps1" echo         // Token check-in
>> "__mk_main.ps1" echo         try
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             _token = await _signIn.LogIn("sergeyr", "rubi69qa2******");
>> "__mk_main.ps1" echo             Console.WriteLine($"Received token in Sign in: {_token}");
>> "__mk_main.ps1" echo         }
>> "__mk_main.ps1" echo         catch (Exception ex)
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             throw new Exception($"Sign-in failed, test is aborted:\n{ex.Message}");
>> "__mk_main.ps1" echo         }
>> "__mk_main.ps1" echo         // Test 1:
>> "__mk_main.ps1" echo         try
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             if (string.IsNullOrEmpty(_token))
>> "__mk_main.ps1" echo             {
>> "__mk_main.ps1" echo 				throw new Exception("Token is null, test is aborted."); 
>> "__mk_main.ps1" echo             }
>> "__mk_main.ps1" echo             else 
>> "__mk_main.ps1" echo 			{
>> "__mk_main.ps1" echo                 string token = await _test1.NextTestStep(_token);
>> "__mk_main.ps1" echo             }
>> "__mk_main.ps1" echo         }
>> "__mk_main.ps1" echo         catch (Exception ex)
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             Console.WriteLine($"Function failed: {ex.Message}");
>> "__mk_main.ps1" echo         }
>> "__mk_main.ps1" echo         _http.Dispose();
>> "__mk_main.ps1" echo     }
>> "__mk_main.ps1" echo }
>> "__mk_main.ps1" echo '@
>> "__mk_main.ps1" echo Set-Content -Path 'MainClass.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_main.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create MainClass.cs & exit /b 1 )
del "__mk_main.ps1"


:: ===== Create SignIn.cs  =====
echo Creating SignIn.cs sample test...

> "__mk_signin.ps1" echo $content = @'
>> "__mk_signin.ps1" echo using NUnit.Framework;
>> "__mk_signin.ps1" echo using System.Net.Http;
>> "__mk_signin.ps1" echo using System.Net.Http.Headers;
>> "__mk_signin.ps1" echo using System.Text;
>> "__mk_signin.ps1" echo using System.Text.Json;
>> "__mk_signin.ps1" echo using static System.Net.WebRequestMethods;
>> "__mk_signin.ps1" echo //
>> "__mk_signin.ps1" echo public class SignIn
>> "__mk_signin.ps1" echo {
>> "__mk_signin.ps1" echo     private readonly HttpClient _httpClient;
>> "__mk_signin.ps1" echo     private readonly JsonSerializerOptions _jsonOps;
>> "__mk_signin.ps1" echo //
>> "__mk_signin.ps1" echo     public SignIn(HttpClient httpClient, JsonSerializerOptions jsonOps) 
>> "__mk_signin.ps1" echo     {
>> "__mk_signin.ps1" echo         _httpClient = httpClient;
>> "__mk_signin.ps1" echo         _jsonOps = jsonOps;
>> "__mk_signin.ps1" echo     }
>> "__mk_signin.ps1" echo     // Sign-in at start of the flow
>> "__mk_signin.ps1" echo     public async Task^<string^> LogIn(string username, string password, CancellationToken cancelToken = default)
>> "__mk_signin.ps1" echo     {
>> "__mk_signin.ps1" echo         var url = "users/v1/signin";
>> "__mk_signin.ps1" echo         var payload = new { username, password };
>> "__mk_signin.ps1" echo         var body = new StringContent(JsonSerializer.Serialize(payload, _jsonOps), Encoding.UTF8, "application/json");
>> "__mk_signin.ps1" echo //
>> "__mk_signin.ps1" echo         using var _responsele = await _httpClient.PostAsync(url, body, cancelToken).ConfigureAwait(false);
>> "__mk_signin.ps1" echo         _responsele.EnsureSuccessStatusCode();
>> "__mk_signin.ps1" echo //
>> "__mk_signin.ps1" echo         var json = await _responsele.Content.ReadAsStringAsync(cancelToken).ConfigureAwait(false);
>> "__mk_signin.ps1" echo         using var doc = JsonDocument.Parse(json);
>> "__mk_signin.ps1" echo         var tokenProp = doc.RootElement.GetProperty("token");
>> "__mk_signin.ps1" echo         var _token = tokenProp.GetString();
>> "__mk_signin.ps1" echo //
>> "__mk_signin.ps1" echo         if (string.IsNullOrEmpty(_token))
>> "__mk_signin.ps1" echo             throw new InvalidOperationException("Token not found in sign-in response.");
>> "__mk_signin.ps1" echo //
>> "__mk_signin.ps1" echo         _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
>> "__mk_signin.ps1" echo         return _token;
>> "__mk_signin.ps1" echo     }
>> "__mk_signin.ps1" echo }
>> "__mk_signin.ps1" echo '@
>> "__mk_signin.ps1" echo Set-Content -Path 'SignIn.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_signin.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create SignIn.cs & exit /b 1 )
del "__mk_signin.ps1"


:: ===== Create Test1.cs  =====
echo Creating Test1.cs sample test...

> "__mk_test1.ps1" echo $content = @'
>> "__mk_test1.ps1" echo public class Test1
>> "__mk_test1.ps1" echo {
>> "__mk_test1.ps1" echo     private readonly HttpClient _client;
>> "__mk_test1.ps1" echo     public Test1 (HttpClient httpClient) 
>> "__mk_test1.ps1" echo     {
>> "__mk_test1.ps1" echo         _client = httpClient;
>> "__mk_test1.ps1" echo     }
>> "__mk_test1.ps1" echo     // Test 1
>> "__mk_test1.ps1" echo     public async Task^<string^> NextTestStep(string token)
>> "__mk_test1.ps1" echo     {
>> "__mk_test1.ps1" echo         // If the token OK?
>> "__mk_test1.ps1" echo         await Task.Run(() =^>
>> "__mk_test1.ps1" echo         {
>> "__mk_test1.ps1" echo             if (string.IsNullOrWhiteSpace(token))
>> "__mk_test1.ps1" echo                 throw new InvalidOperationException("Token is null/empty. Call SignIn first.");
>> "__mk_test1.ps1" echo         });
>> "__mk_test1.ps1" echo         Console.WriteLine($"Received token in the Next Test Step:\n{token}");
>> "__mk_test1.ps1" echo         return token;
>> "__mk_test1.ps1" echo     }
>> "__mk_test1.ps1" echo }
>> "__mk_test1.ps1" echo '@
>> "__mk_test1.ps1" echo Set-Content -Path 'Test1.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_test1.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create Test1.cs & exit /b 1 )
del "__mk_test1.ps1"

:: ===== Create WebDriverSettings.cs  =====
echo Creating WebDriverSettings.cs sample test...

> "__mk_webdriver.ps1" echo $content = @'
>> "__mk_webdriver.ps1" echo using OpenQA.Selenium;
>> "__mk_webdriver.ps1" echo using OpenQA.Selenium.Chrome;
>> "__mk_webdriver.ps1" echo using Microsoft.Data.SqlClient;
>> "__mk_webdriver.ps1" echo using System;
>> "__mk_webdriver.ps1" echo //
>> "__mk_webdriver.ps1" echo namespace Test.GlobalClasses
>> "__mk_webdriver.ps1" echo {
>> "__mk_webdriver.ps1" echo     public class WebDriverSettings
>> "__mk_webdriver.ps1" echo     {
>> "__mk_webdriver.ps1" echo         private static IWebDriver _webDriver;
>> "__mk_webdriver.ps1" echo         private static readonly object _lockObject = new object();
>> "__mk_webdriver.ps1" echo //
>> "__mk_webdriver.ps1" echo         public static IWebDriver WebDriver
>> "__mk_webdriver.ps1" echo         {
>> "__mk_webdriver.ps1" echo             get
>> "__mk_webdriver.ps1" echo             { 
>> "__mk_webdriver.ps1" echo                 if (_webDriver == null)
>> "__mk_webdriver.ps1" echo                 {
>> "__mk_webdriver.ps1" echo                     lock (_lockObject)
>> "__mk_webdriver.ps1" echo                     {
>> "__mk_webdriver.ps1" echo                         if (_webDriver == null)
>> "__mk_webdriver.ps1" echo                         {
>> "__mk_webdriver.ps1" echo                             _webDriver = new ChromeDriver();
>> "__mk_webdriver.ps1" echo                             _webDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
>> "__mk_webdriver.ps1" echo                         }
>> "__mk_webdriver.ps1" echo                     }
>> "__mk_webdriver.ps1" echo                 }
>> "__mk_webdriver.ps1" echo                 return _webDriver;
>> "__mk_webdriver.ps1" echo             } //get
>> "__mk_webdriver.ps1" echo         }
>> "__mk_webdriver.ps1" echo //
>> "__mk_webdriver.ps1" echo         public static void DisposeWebDriver()
>> "__mk_webdriver.ps1" echo         {
>> "__mk_webdriver.ps1" echo             _webDriver?.Quit();
>> "__mk_webdriver.ps1" echo             _webDriver = null;
>> "__mk_webdriver.ps1" echo         }
>> "__mk_webdriver.ps1" echo     }
>> "__mk_webdriver.ps1" echo }
>> "__mk_webdriver.ps1" echo '@
>> "__mk_webdriver.ps1" echo Set-Content -Path 'WebDriverSettings.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_webdriver.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create WebDriverSettings.cs & exit /b 1 )
del "__mk_webdriver.ps1"

:: ===== Create SqlOperations.cs — temp PS1, column-1 here-string =====
echo Creating SqlOperations.cs...

> "__mk_sqlops.ps1" echo $content = @'
>> "__mk_sqlops.ps1" echo using Microsoft.Data.SqlClient;
>> "__mk_sqlops.ps1" echo //
>> "__mk_sqlops.ps1" echo namespace SqlOperations
>> "__mk_sqlops.ps1" echo {
>> "__mk_sqlops.ps1" echo     class DbData
>> "__mk_sqlops.ps1" echo     {
>> "__mk_sqlops.ps1" echo         private readonly string ConnectionToDBA = "Server=qa2ilsql01;Database=DCS;Integrated Security=True";
>> "__mk_sqlops.ps1" echo //
>> "__mk_sqlops.ps1" echo         public async Task^<int^> DeleteFromDB()
>> "__mk_sqlops.ps1" echo         {
>> "__mk_sqlops.ps1" echo             string query = @"
>> "__mk_sqlops.ps1" echo USE DCS_SVC_SCHEDULING;
>> "__mk_sqlops.ps1" echo DECLARE @EntityId    BIGINT = 000000000;
>> "__mk_sqlops.ps1" echo DECLARE @CountBefore INT, @DeletedRows INT;
>> "__mk_sqlops.ps1" echo BEGIN TRAN delete_tasks;
>> "__mk_sqlops.ps1" echo SELECT @CountBefore = COUNT(*)
>> "__mk_sqlops.ps1" echo FROM dbo.TBL_NAME WITH (UPDLOCK, HOLDLOCK)
>> "__mk_sqlops.ps1" echo WHERE col_name = @EntityId;
>> "__mk_sqlops.ps1" echo DELETE dbo.TBL_NAME
>> "__mk_sqlops.ps1" echo WHERE col_name = @EntityId;
>> "__mk_sqlops.ps1" echo SET @DeletedRows = @@ROWCOUNT;
>> "__mk_sqlops.ps1" echo IF (@DeletedRows ^> @CountBefore)
>> "__mk_sqlops.ps1" echo BEGIN
>> "__mk_sqlops.ps1" echo     ROLLBACK TRAN delete_tasks;
>> "__mk_sqlops.ps1" echo     RAISERROR('Too many records to delete. Count=%%%d, Deleted=%%%d', 16, 1, @CountBefore, @DeletedRows);
>> "__mk_sqlops.ps1" echo END
>> "__mk_sqlops.ps1" echo ELSE
>> "__mk_sqlops.ps1" echo BEGIN
>> "__mk_sqlops.ps1" echo     COMMIT TRAN delete_tasks;
>> "__mk_sqlops.ps1" echo     PRINT CAST(@DeletedRows AS VARCHAR(33)) + ' record(s) have been deleted.';
>> "__mk_sqlops.ps1" echo END";
>> "__mk_sqlops.ps1" echo             using (SqlConnection conn = new SqlConnection(ConnectionToDBA))
>> "__mk_sqlops.ps1" echo             {
>> "__mk_sqlops.ps1" echo                 try { await conn.OpenAsync(); }
>> "__mk_sqlops.ps1" echo                 catch (SqlException ex)
>> "__mk_sqlops.ps1" echo                 {
>> "__mk_sqlops.ps1" echo                     Console.WriteLine($"SQL Exception: {ex.Message}");
>> "__mk_sqlops.ps1" echo                     return -1;
>> "__mk_sqlops.ps1" echo                 }
>> "__mk_sqlops.ps1" echo                 using (SqlCommand cmd = new SqlCommand(query, conn))
>> "__mk_sqlops.ps1" echo                 {
>> "__mk_sqlops.ps1" echo                     int rowsAffected = await cmd.ExecuteNonQueryAsync();
>> "__mk_sqlops.ps1" echo                     Console.WriteLine($"Rows affected (pre-test cleanup): {rowsAffected}");
>> "__mk_sqlops.ps1" echo                     return rowsAffected;
>> "__mk_sqlops.ps1" echo                 }
>> "__mk_sqlops.ps1" echo             }
>> "__mk_sqlops.ps1" echo         }
>> "__mk_sqlops.ps1" echo     }
>> "__mk_sqlops.ps1" echo }
>> "__mk_sqlops.ps1" echo '@
>> "__mk_sqlops.ps1" echo Set-Content -Path 'SqlOperations.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_sqlops.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create SqlOperations.cs & exit /b 1 )
del "__mk_sqlops.ps1"


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
