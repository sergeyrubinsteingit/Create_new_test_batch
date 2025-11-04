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
set "DEFAULT_PATH=C:\Users\sergeyr\WEB_QA_TESTS"
set /p BASE_PATH=Enter the full path where you want to create a new directory (default %DEFAULT_PATH%): 
if "%BASE_PATH%"=="" set "BASE_PATH=%DEFAULT_PATH%"



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



:: ===== Create Global Classes dir =====
set "GLOBAL_CLASSES_DIR=%FULL_PROJECT_PATH%\GlobalClasses"
if not exist %GLOBAL_CLASSES_DIR% (
	mkdir %GLOBAL_CLASSES_DIR% || ( echo [ERROR] Failed to create "%GLOBAL_CLASSES_DIR%". & exit /b 1)
)



:: ===== Create Scheduling dir =====
set "SCHEDULING_DIR=%FULL_PROJECT_PATH%\Scheduling"
if not exist "%SCHEDULING_DIR%" (
	mkdir %SCHEDULING_DIR% || ( echo [ERROR] Failed to create "%SCHEDULING_DIR%".& exit /b 1 )
)



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
dotnet add "%CSPROJ%" package Microsoft.Data.SqlClient
dotnet add "%CSPROJ%" package Twilio
dotnet add "%CSPROJ%" package Google.Apis.Gmail.v1
dotnet add "%CSPROJ%" package Google.Apis.Auth
dotnet add "%CSPROJ%" package Microsoft.Office.Interop.Outlook
dotnet add "%CSPROJ%" package Microsoft.Windows.Compatibility



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
>> "__mk_main.ps1" echo using GlobalClasses;
>> "__mk_main.ps1" echo using System.Text.Json;
>> "__mk_main.ps1" echo using HttpClient = System.Net.Http.HttpClient;
>> "__mk_main.ps1" echo //
>> "__mk_main.ps1" echo public class MainClass
>> "__mk_main.ps1" echo {
>> "__mk_main.ps1" echo     private readonly HttpClient _http;
>> "__mk_main.ps1" echo     private readonly JsonSerializerOptions _json = new(JsonSerializerDefaults.Web);
>> "__mk_main.ps1" echo     private readonly GlobalVariables _globals = new();
>> "__mk_main.ps1" echo     private readonly Test1 _test1;
>> "__mk_main.ps1" echo //
>> "__mk_main.ps1" echo     public MainClass(HttpClient? httpClient = null)
>> "__mk_main.ps1" echo     {
>> "__mk_main.ps1" echo         _http = httpClient ?? new HttpClient();
>> "__mk_main.ps1" echo         _test1 = new Test1(_http);
>> "__mk_main.ps1" echo     }
>> "__mk_main.ps1" echo     public static async Task Main(string[] args)
>> "__mk_main.ps1" echo     {
>> "__mk_main.ps1" echo         var p = new MainClass();
>> "__mk_main.ps1" echo         await p.RunAsync();
>> "__mk_main.ps1" echo     }
>> "__mk_main.ps1" echo     public async Task RunAsync()
>> "__mk_main.ps1" echo     {
>> "__mk_main.ps1" echo         try
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             // One login here:
>> "__mk_main.ps1" echo             _http.BaseAddress = new Uri(_globals.BaseUrl);
>> "__mk_main.ps1" echo             await _globals.InitializeAsync(_http, _json); // sets TokenKoken and Authorization header
>> "__mk_main.ps1" echo             Console.WriteLine($"Token received: {_globals.TokenKoken[..Math.Min(12, _globals.TokenKoken.Length)]}...");
>> "__mk_main.ps1" echo         }
>> "__mk_main.ps1" echo         catch (Exception ex)
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             throw new Exception($"Sign-in failed, test is aborted:\n{ex.Message}");
>> "__mk_main.ps1" echo         }
>> "__mk_main.ps1" echo         try
>> "__mk_main.ps1" echo         {
>> "__mk_main.ps1" echo             // use the same HttpClient which already has Bearer header
>> "__mk_main.ps1" echo             string result = await _test1.NextTestStep(_globals.TokenKoken);
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
>> "__mk_signin.ps1" echo using System.Net.Http.Headers;
>> "__mk_signin.ps1" echo using System.Text;
>> "__mk_signin.ps1" echo using System.Text.Json;
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
>> "__mk_signin.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/SignIn.cs' -Value $content -Encoding UTF8

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
>> "__mk_test1.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/Test1.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_test1.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create Test1.cs & exit /b 1 )
del "__mk_test1.ps1"



:: ===== Create WebDriverSettings.cs  =====
echo Creating WebDriverSettings.cs sample test...

> "__mk_webdriver.ps1" echo $content = @'
>> "__mk_webdriver.ps1" echo using OpenQA.Selenium;
>> "__mk_webdriver.ps1" echo using OpenQA.Selenium.Chrome;
>> "__mk_webdriver.ps1" echo //
>> "__mk_webdriver.ps1" echo namespace GlobalClasses
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
>> "__mk_webdriver.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/WebDriverSettings.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_webdriver.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create WebDriverSettings.cs & exit /b 1 )
del "__mk_webdriver.ps1"



:: ===== Create SqlOperations.cs =====
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
>> "__mk_sqlops.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/SqlOperations.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_sqlops.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create SqlOperations.cs & exit /b 1 )
del "__mk_sqlops.ps1"



:: ===== Create OAthClientSecret.json =====
echo Creating OAthClientSecret.json...

> "__mk_oauth_secret.ps1" echo $content = @'
>> "__mk_oauth_secret.ps1" echo {
>> "__mk_oauth_secret.ps1" echo     "web": {
>> "__mk_oauth_secret.ps1" echo         "client_id": "846414882058-fcv59p32enos76090akmv2uhjg4rkbvh.apps.googleusercontent.com",
>> "__mk_oauth_secret.ps1" echo         "project_id": "postman-gmail-test-470312",
>> "__mk_oauth_secret.ps1" echo         "auth_uri": "https://accounts.google.com/o/oauth2/auth",
>> "__mk_oauth_secret.ps1" echo         "token_uri": "https://oauth2.googleapis.com/token",
>> "__mk_oauth_secret.ps1" echo         "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
>> "__mk_oauth_secret.ps1" echo         "client_secret": "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",
>> "__mk_oauth_secret.ps1" echo         "redirect_uris": [
>> "__mk_oauth_secret.ps1" echo             "https://oauth.pstmn.io/v1/callback"
>> "__mk_oauth_secret.ps1" echo         ]
>> "__mk_oauth_secret.ps1" echo     }
>> "__mk_oauth_secret.ps1" echo }
>> "__mk_oauth_secret.ps1" echo '@
>> "__mk_oauth_secret.ps1" echo Set-Content -Path 'OAthClientSecret.json' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_oauth_secret.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create OAthClientSecret.json & exit /b 1 )
del "__mk_oauth_secret.ps1"



:: ===== Create OutlookMessagesHandler.cs =====
echo Creating OutlookMessagesHandler.cs...

> "__mk_outlook_msgs.ps1" echo $content = @'
>> "__mk_outlook_msgs.ps1" echo using OpenQA.Selenium;
>> "__mk_outlook_msgs.ps1" echo using System.Diagnostics;
>> "__mk_outlook_msgs.ps1" echo using System.Reflection;
>> "__mk_outlook_msgs.ps1" echo using System.Runtime.InteropServices;
>> "__mk_outlook_msgs.ps1" echo using System.Text.RegularExpressions;
>> "__mk_outlook_msgs.ps1" echo using Outlook_ = Microsoft.Office.Interop.Outlook;
>> "__mk_outlook_msgs.ps1" echo //
>> "__mk_outlook_msgs.ps1" echo namespace GlobalClasses
>> "__mk_outlook_msgs.ps1" echo {
>> "__mk_outlook_msgs.ps1" echo     class OutlookMessagesHandler : WebDriverSettings
>> "__mk_outlook_msgs.ps1" echo     {
>> "__mk_outlook_msgs.ps1" echo         private static readonly IWebDriver _driver = WebDriver;
>> "__mk_outlook_msgs.ps1" echo         public static string LoginUrl = "";
>> "__mk_outlook_msgs.ps1" echo         public static long maxAttempts = 10000000;
>> "__mk_outlook_msgs.ps1" echo 		//
>> "__mk_outlook_msgs.ps1" echo         public static Outlook_.Application ProceedInvitationLink(bool isDetailed)
>> "__mk_outlook_msgs.ps1" echo         {
>> "__mk_outlook_msgs.ps1" echo             try
>> "__mk_outlook_msgs.ps1" echo             {
>> "__mk_outlook_msgs.ps1" echo                 Console.WriteLine("\n<<<<<<<<<<< Proceed Invitation Link method begins >>>>>>>>>>>>>>>\n");
>> "__mk_outlook_msgs.ps1" echo                 // Create the Outlook application, in-line initialization
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.Application OutlookApp = new Outlook_.Application();
>> "__mk_outlook_msgs.ps1" echo                 // Get the MAPI namespace
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.NameSpace OutlookNameSpace = OutlookApp.GetNamespace("mapi");
>> "__mk_outlook_msgs.ps1" echo                 // Log on by using the default profile or existing session (no dialog box)
>> "__mk_outlook_msgs.ps1" echo                 OutlookNameSpace.Logon(Missing.Value, Missing.Value, false, true);
>> "__mk_outlook_msgs.ps1" echo                 // Get the Inbox folder ^> Nayax_User_Invitations subfolder
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.MAPIFolder InboxSubfolder = OutlookNameSpace.GetDefaultFolder(Outlook_.OlDefaultFolders.olFolderInbox).Folders["Nayax_User_Invitations"];
>> "__mk_outlook_msgs.ps1" echo                 // Get the Items collection in the Inbox folder.
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.Items SubfolderItems;
>> "__mk_outlook_msgs.ps1" echo                 // Get the first message by subject
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.MailItem FirstMessage = null;
>> "__mk_outlook_msgs.ps1" echo                 // Waiting for the first message // System.NullReferenceException:
>> "__mk_outlook_msgs.ps1" echo                 int _attempt = 1;
>> "__mk_outlook_msgs.ps1" echo                 do
>> "__mk_outlook_msgs.ps1" echo                 {
>> "__mk_outlook_msgs.ps1" echo                     try
>> "__mk_outlook_msgs.ps1" echo                     {
>> "__mk_outlook_msgs.ps1" echo                         // forced pause
>> "__mk_outlook_msgs.ps1" echo                         Thread.Sleep(1000);
>> "__mk_outlook_msgs.ps1" echo                         // Refresh the Items collection
>> "__mk_outlook_msgs.ps1" echo                         SubfolderItems = InboxSubfolder.Items;
>> "__mk_outlook_msgs.ps1" echo                         // Get the first message by subject
>> "__mk_outlook_msgs.ps1" echo                         FirstMessage = (Outlook_.MailItem)SubfolderItems.Find("[Subject] = Nayax User Invitation");
>> "__mk_outlook_msgs.ps1" echo                         // If the message is found, break the loop
>> "__mk_outlook_msgs.ps1" echo                         if (FirstMessage == null)
>> "__mk_outlook_msgs.ps1" echo                         {
>> "__mk_outlook_msgs.ps1" echo                             Console.WriteLine("First message found on attempt [ " + _attempt + " ]");
>> "__mk_outlook_msgs.ps1" echo                             break;
>> "__mk_outlook_msgs.ps1" echo                         }//if
>> "__mk_outlook_msgs.ps1" echo                         // If the message is not finally found
>> "__mk_outlook_msgs.ps1" echo                         if (_attempt ^> maxAttempts)
>> "__mk_outlook_msgs.ps1" echo                         {
>> "__mk_outlook_msgs.ps1" echo                             // exit
>> "__mk_outlook_msgs.ps1" echo                             _driver.Quit();
>> "__mk_outlook_msgs.ps1" echo                             throw new Exception("The first message was not found after " + maxAttempts + " attempts. The test is shut down.");
>> "__mk_outlook_msgs.ps1" echo                         }//if
>> "__mk_outlook_msgs.ps1" echo                     }
>> "__mk_outlook_msgs.ps1" echo                     catch (Exception)
>> "__mk_outlook_msgs.ps1" echo                     {
>> "__mk_outlook_msgs.ps1" echo                         _attempt++;
>> "__mk_outlook_msgs.ps1" echo                         Console.WriteLine("Waiting for the first message, attempt [ " + _attempt.ToString() + " ]");
>> "__mk_outlook_msgs.ps1" echo                         // forced pause
>> "__mk_outlook_msgs.ps1" echo                         Thread.Sleep(BandwidthCheck.DownloadRate = Convert.ToInt32(BandwidthCheck.RunBandwidthCheckAsync()));
>> "__mk_outlook_msgs.ps1" echo                         continue;
>> "__mk_outlook_msgs.ps1" echo                     }//catch
>> "__mk_outlook_msgs.ps1" echo                 } while (FirstMessage == null);
>> "__mk_outlook_msgs.ps1" echo                 LoginUrl = Regex.Match(FirstMessage.Body.ToString(), @"Log in now <(.+?)>").Groups[1].Value;
>> "__mk_outlook_msgs.ps1" echo                 Console.WriteLine("User registration URL:\n" + LoginUrl + "\n");
>> "__mk_outlook_msgs.ps1" echo                 // just in case we want to see some mail details
>> "__mk_outlook_msgs.ps1" echo                 if (isDetailed)
>> "__mk_outlook_msgs.ps1" echo                 {
>> "__mk_outlook_msgs.ps1" echo                     //Check for attachments
>> "__mk_outlook_msgs.ps1" echo                     int AttachCnt = FirstMessage.Attachments.Count;
>> "__mk_outlook_msgs.ps1" echo                     Console.WriteLine("Attachments: " + AttachCnt.ToString());
>> "__mk_outlook_msgs.ps1" echo                     //some common properties.
>> "__mk_outlook_msgs.ps1" echo                     Console.WriteLine(FirstMessage.Subject);
>> "__mk_outlook_msgs.ps1" echo                     Console.WriteLine(FirstMessage.SenderName);
>> "__mk_outlook_msgs.ps1" echo                     Console.WriteLine(FirstMessage.ReceivedTime);
>> "__mk_outlook_msgs.ps1" echo                     //Attachments' listing
>> "__mk_outlook_msgs.ps1" echo                     if (AttachCnt ^> 0)
>> "__mk_outlook_msgs.ps1" echo                     {
>> "__mk_outlook_msgs.ps1" echo                         for (int i = 1; i ^<= AttachCnt; i++) Console.WriteLine(i.ToString() + "-" + FirstMessage.Attachments[i].DisplayName);
>> "__mk_outlook_msgs.ps1" echo                     }//if
>> "__mk_outlook_msgs.ps1" echo                     //Display the Outlook window
>> "__mk_outlook_msgs.ps1" echo                     FirstMessage.Display(true); //modal
>> "__mk_outlook_msgs.ps1" echo                 };//if
>> "__mk_outlook_msgs.ps1" echo                 //Log off.
>> "__mk_outlook_msgs.ps1" echo                 OutlookNameSpace.Logoff();
>> "__mk_outlook_msgs.ps1" echo                 //Explicitly release objects.
>> "__mk_outlook_msgs.ps1" echo                 FirstMessage = null;
>> "__mk_outlook_msgs.ps1" echo                 SubfolderItems = null;
>> "__mk_outlook_msgs.ps1" echo                 InboxSubfolder = null;
>> "__mk_outlook_msgs.ps1" echo                 OutlookNameSpace = null;
>> "__mk_outlook_msgs.ps1" echo                 OutlookApp = null;
>> "__mk_outlook_msgs.ps1" echo                 // If everything is OK, the Outlook application is returned.
>> "__mk_outlook_msgs.ps1" echo                 return OutlookApp;
>> "__mk_outlook_msgs.ps1" echo             }
>> "__mk_outlook_msgs.ps1" echo             //Error handler.
>> "__mk_outlook_msgs.ps1" echo             catch (Exception ex)
>> "__mk_outlook_msgs.ps1" echo             {
>> "__mk_outlook_msgs.ps1" echo                 _driver.Quit();
>> "__mk_outlook_msgs.ps1" echo                 throw new Exception("ProceedInvitationLink failed. Trace:\n", ex);
>> "__mk_outlook_msgs.ps1" echo             }
>> "__mk_outlook_msgs.ps1" echo         }
>> "__mk_outlook_msgs.ps1" echo 		//
>> "__mk_outlook_msgs.ps1" echo         public static Outlook_.Application DeleteExistingMails()
>> "__mk_outlook_msgs.ps1" echo         {
>> "__mk_outlook_msgs.ps1" echo             Console.WriteLine("\n<<<<<<<<<<< DeleteExistingMails method begins >>>>>>>>>>>>>>>\n");
>> "__mk_outlook_msgs.ps1" echo             try
>> "__mk_outlook_msgs.ps1" echo             {
>> "__mk_outlook_msgs.ps1" echo                 // Create the Outlook application, in-line initialization
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.Application OutlookApp = new Outlook_.Application();
>> "__mk_outlook_msgs.ps1" echo                 // Get the MAPI namespace
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.NameSpace OutlookNameSpace = OutlookApp.GetNamespace("mapi");
>> "__mk_outlook_msgs.ps1" echo                 // Log on by using the default profile or existing session (no dialog box)
>> "__mk_outlook_msgs.ps1" echo                 OutlookNameSpace.Logon(Missing.Value, Missing.Value, false, true);
>> "__mk_outlook_msgs.ps1" echo                 // Get the Inbox folder ^> Nayax_User_Invitations subfolder
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.MAPIFolder InboxSubfolder = OutlookNameSpace.GetDefaultFolder(Outlook_.OlDefaultFolders.olFolderInbox).Folders["Nayax_User_Invitations"];
>> "__mk_outlook_msgs.ps1" echo                 // Get the Items collection in the Inbox folder.
>> "__mk_outlook_msgs.ps1" echo                 Outlook_.Items SubfolderItems = InboxSubfolder.Items;
>> "__mk_outlook_msgs.ps1" echo                 Console.WriteLine($"SubfolderItems _attempt: {SubfolderItems.Count.ToString()}");
>> "__mk_outlook_msgs.ps1" echo                 if (SubfolderItems.Count ^> 0)
>> "__mk_outlook_msgs.ps1" echo                 {
>> "__mk_outlook_msgs.ps1" echo                     for (int i = SubfolderItems.Count; i ^>= 1; i--)
>> "__mk_outlook_msgs.ps1" echo                     {
>> "__mk_outlook_msgs.ps1" echo                         var item = SubfolderItems[i] as Outlook_.MailItem;
>> "__mk_outlook_msgs.ps1" echo                         if (item == null)
>> "__mk_outlook_msgs.ps1" echo                         {
>> "__mk_outlook_msgs.ps1" echo                             Console.WriteLine("Deleting message: " + item.Subject);
>> "__mk_outlook_msgs.ps1" echo                             item.Delete();
>> "__mk_outlook_msgs.ps1" echo                         }
>> "__mk_outlook_msgs.ps1" echo                     }
>> "__mk_outlook_msgs.ps1" echo                 }
>> "__mk_outlook_msgs.ps1" echo                 //Log off.
>> "__mk_outlook_msgs.ps1" echo                 OutlookNameSpace.Logoff();
>> "__mk_outlook_msgs.ps1" echo                 //Explicitly release objects.
>> "__mk_outlook_msgs.ps1" echo                 //FirstMessage = null;
>> "__mk_outlook_msgs.ps1" echo                 SubfolderItems = null;
>> "__mk_outlook_msgs.ps1" echo                 InboxSubfolder = null;
>> "__mk_outlook_msgs.ps1" echo                 OutlookNameSpace = null;
>> "__mk_outlook_msgs.ps1" echo                 OutlookApp = null;
>> "__mk_outlook_msgs.ps1" echo                 // If everything is OK, the Outlook application is returned.
>> "__mk_outlook_msgs.ps1" echo                 return OutlookApp;
>> "__mk_outlook_msgs.ps1" echo             }
>> "__mk_outlook_msgs.ps1" echo             catch (Exception e) when (e is COMException)
>> "__mk_outlook_msgs.ps1" echo             {
>> "__mk_outlook_msgs.ps1" echo                 _driver.Quit();
>> "__mk_outlook_msgs.ps1" echo                 throw new Exception($"DeleteExistingMails failed. Trace:\n{e.StackTrace}");
>> "__mk_outlook_msgs.ps1" echo             }//catch
>> "__mk_outlook_msgs.ps1" echo         }//DeleteExistingMails
>> "__mk_outlook_msgs.ps1" echo 		//
>> "__mk_outlook_msgs.ps1" echo         public static Outlook_.Application CloseOutlookIfOpened()
>> "__mk_outlook_msgs.ps1" echo         {
>> "__mk_outlook_msgs.ps1" echo             Console.WriteLine("\n<<<<<<<<< CloseOutlook begins >>>>>>>>>\n");
>> "__mk_outlook_msgs.ps1" echo             bool isOutlookClosed = false;
>> "__mk_outlook_msgs.ps1" echo             try
>> "__mk_outlook_msgs.ps1" echo             {
>> "__mk_outlook_msgs.ps1" echo                 var outlookApp = new Outlook_.Application();
>> "__mk_outlook_msgs.ps1" echo                 outlookApp.Quit();
>> "__mk_outlook_msgs.ps1" echo                 Marshal.ReleaseComObject(outlookApp);
>> "__mk_outlook_msgs.ps1" echo                 GC.Collect();
>> "__mk_outlook_msgs.ps1" echo                 GC.WaitForPendingFinalizers();
>> "__mk_outlook_msgs.ps1" echo                 // forced pause
>> "__mk_outlook_msgs.ps1" echo                 Thread.Sleep(Convert.ToInt32(BandwidthCheck.DownloadRate = Convert.ToInt32(BandwidthCheck.RunBandwidthCheckAsync()) * 2 * 10));
>> "__mk_outlook_msgs.ps1" echo                 if (Process.GetProcessesByName("OUTLOOK").Length == 0) isOutlookClosed = true;
>> "__mk_outlook_msgs.ps1" echo             }
>> "__mk_outlook_msgs.ps1" echo             catch
>> "__mk_outlook_msgs.ps1" echo             {
>> "__mk_outlook_msgs.ps1" echo                 // Outlook isn't running
>> "__mk_outlook_msgs.ps1" echo                 isOutlookClosed = true;
>> "__mk_outlook_msgs.ps1" echo             }//catch
>> "__mk_outlook_msgs.ps1" echo             if (isOutlookClosed) 
>> "__mk_outlook_msgs.ps1" echo             { 
>> "__mk_outlook_msgs.ps1" echo                 foreach (var process in Process.GetProcessesByName("OUTLOOK")) 
>> "__mk_outlook_msgs.ps1" echo                 { 
>> "__mk_outlook_msgs.ps1" echo                     process.Kill();
>> "__mk_outlook_msgs.ps1" echo                     process.WaitForExit();
>> "__mk_outlook_msgs.ps1" echo                 }//foreach
>> "__mk_outlook_msgs.ps1" echo             }//if
>> "__mk_outlook_msgs.ps1" echo             return null;
>> "__mk_outlook_msgs.ps1" echo         }// CloseOutlookIfOpened
>> "__mk_outlook_msgs.ps1" echo     }
>> "__mk_outlook_msgs.ps1" echo }
>> "__mk_outlook_msgs.ps1" echo '@
>> "__mk_outlook_msgs.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/OutlookMessagesHandler.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_outlook_msgs.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create OutlookMessagesHandler.cs & exit /b 1 )
del "__mk_outlook_msgs.ps1"



:: ===== Create BandwidthCheck.cs =====
echo Creating BandwidthCheck.cs...

> "__mk_bandwidth.ps1" echo $content = @'
>> "__mk_bandwidth.ps1" echo using System.Diagnostics;
>> "__mk_bandwidth.ps1" echo using System.Net.Http.Headers;
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo namespace GlobalClasses
>> "__mk_bandwidth.ps1" echo {
>> "__mk_bandwidth.ps1" echo     public class BandwidthCheck : WebDriverSettings
>> "__mk_bandwidth.ps1" echo     {
>> "__mk_bandwidth.ps1" echo         public static int DownloadRate = 0;
>> "__mk_bandwidth.ps1" echo         public static double Coefficient_ = 0;
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo         private static readonly HttpClient _http = new HttpClient();
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo         public static async Task^<int^> RunBandwidthCheckAsync(CancellationToken cancellationToken = default)
>> "__mk_bandwidth.ps1" echo         {
>> "__mk_bandwidth.ps1" echo             Console.WriteLine("\n<<<<<<<<< BandwidthCheck method begins (async) >>>>>>>>>\n");
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo             const string url = "https://media.geeksforgeeks.org/wp-content/uploads/gfg-40.png";
>> "__mk_bandwidth.ps1" echo             const string outFile = "gfg-40.png";
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo             for (var i = 0; i ^< 10; i++)
>> "__mk_bandwidth.ps1" echo             {
>> "__mk_bandwidth.ps1" echo                 try
>> "__mk_bandwidth.ps1" echo                 {
>> "__mk_bandwidth.ps1" echo                     // Forsed pause
>> "__mk_bandwidth.ps1" echo                     await Task.Delay(5, cancellationToken);
>> "__mk_bandwidth.ps1" echo                     // Per-iteration timing
>> "__mk_bandwidth.ps1" echo                     var sw = Stopwatch.StartNew();
>> "__mk_bandwidth.ps1" echo                     // Cache-buster to avoid local/proxy cache
>> "__mk_bandwidth.ps1" echo                     var bust = $"?nocache={Guid.NewGuid():N}";
>> "__mk_bandwidth.ps1" echo                     using var req = new HttpRequestMessage(HttpMethod.Get, url + bust);
>> "__mk_bandwidth.ps1" echo                     req.Headers.CacheControl = new CacheControlHeaderValue { NoCache = true };
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo                     using var response = await _http.SendAsync(
>> "__mk_bandwidth.ps1" echo                         req,
>> "__mk_bandwidth.ps1" echo                         HttpCompletionOption.ResponseHeadersRead,
>> "__mk_bandwidth.ps1" echo                         cancellationToken);
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo                     response.EnsureSuccessStatusCode();
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo                     await using var httpStream = await response.Content.ReadAsStreamAsync(cancellationToken);
>> "__mk_bandwidth.ps1" echo                     await using var fileStream = File.Create(outFile);
>> "__mk_bandwidth.ps1" echo                     await httpStream.CopyToAsync(fileStream, cancellationToken);
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo                     sw.Stop();
>> "__mk_bandwidth.ps1" echo                     DownloadRate = Convert.ToInt32(sw.Elapsed.TotalMilliseconds);
>> "__mk_bandwidth.ps1" echo                 }
>> "__mk_bandwidth.ps1" echo                 catch (Exception)
>> "__mk_bandwidth.ps1" echo                 {
>> "__mk_bandwidth.ps1" echo                     Console.WriteLine("Image is not found to measure downloading rate: \"gfg-40.png\"");
>> "__mk_bandwidth.ps1" echo                     await Task.Delay(5, cancellationToken);
>> "__mk_bandwidth.ps1" echo                 }
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo                 await Task.Delay(50, cancellationToken);
>> "__mk_bandwidth.ps1" echo             }
>> "__mk_bandwidth.ps1" echo             // Coefficient in case of resault being too large
>> "__mk_bandwidth.ps1" echo             Coefficient_ = DownloadRate ^>= 1000 ? 0.1 : 100;
>> "__mk_bandwidth.ps1" echo //
>> "__mk_bandwidth.ps1" echo             Console.WriteLine($">>>>>>>>>> DownloadRate: {DownloadRate} ms. <<<<<<<<<<");
>> "__mk_bandwidth.ps1" echo             return DownloadRate;
>> "__mk_bandwidth.ps1" echo         }
>> "__mk_bandwidth.ps1" echo     }
>> "__mk_bandwidth.ps1" echo }
>> "__mk_bandwidth.ps1" echo '@
>> "__mk_bandwidth.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/BandwidthCheck.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_bandwidth.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create Create BandwidthCheck.cs & exit /b 1 )
del "__mk_bandwidth.ps1"



:: ===== Create Helpers.cs =====
echo Creating Helpers.cs...

> "__mk_helpers.ps1" echo $content = @'
>> "__mk_helpers.ps1" echo using System.Globalization;
>> "__mk_helpers.ps1" echo using System.Net.Http.Headers;
>> "__mk_helpers.ps1" echo //
>> "__mk_helpers.ps1" echo namespace GlobalClasses
>> "__mk_helpers.ps1" echo {
>> "__mk_helpers.ps1" echo     public class Helpers
>> "__mk_helpers.ps1" echo     {
>> "__mk_helpers.ps1" echo         // =========================== Assert and Http ===========================
>> "__mk_helpers.ps1" echo         public static class AssertEx
>> "__mk_helpers.ps1" echo         {
>> "__mk_helpers.ps1" echo             public static void Require(bool condition, string messageIfFails)
>> "__mk_helpers.ps1" echo             {
>> "__mk_helpers.ps1" echo                 if (condition) throw new Exception("[ASSERT] " + messageIfFails);
>> "__mk_helpers.ps1" echo                 Console.WriteLine("[OK] " + messageIfFails);
>> "__mk_helpers.ps1" echo             }
>> "__mk_helpers.ps1" echo         }
>> "__mk_helpers.ps1" echo         public static class HttpEx
>> "__mk_helpers.ps1" echo         {
>> "__mk_helpers.ps1" echo             public static readonly HttpClient Client = new()
>> "__mk_helpers.ps1" echo             {
>> "__mk_helpers.ps1" echo                 Timeout = TimeSpan.FromSeconds(100)
>> "__mk_helpers.ps1" echo             };
>> "__mk_helpers.ps1" echo             public static void EnsureBearer(GlobalVariables _context)
>> "__mk_helpers.ps1" echo             {
>> "__mk_helpers.ps1" echo                 {
>> "__mk_helpers.ps1" echo                     Client.DefaultRequestHeaders.Authorization =
>> "__mk_helpers.ps1" echo                         new AuthenticationHeaderValue("Bearer", _context.TokenKoken);
>> "__mk_helpers.ps1" echo                 }
>> "__mk_helpers.ps1" echo             }
>> "__mk_helpers.ps1" echo //
>> "__mk_helpers.ps1" echo             public static Task SleepDelay(int ms, CancellationToken cancelToken = default) =^>
>> "__mk_helpers.ps1" echo                 Task.Delay(ms, cancelToken);
>> "__mk_helpers.ps1" echo         }
>> "__mk_helpers.ps1" echo     }
>> "__mk_helpers.ps1" echo // =========================== TimeUtil ===========================
>> "__mk_helpers.ps1" echo     public static class TimeUtil
>> "__mk_helpers.ps1" echo     {
>> "__mk_helpers.ps1" echo         private static string? TryGetTimeZone(string id)
>> "__mk_helpers.ps1" echo         {
>> "__mk_helpers.ps1" echo             try { _ = TimeZoneInfo.FindSystemTimeZoneById(id); return id; }
>> "__mk_helpers.ps1" echo             catch { return null; }
>> "__mk_helpers.ps1" echo         }
>> "__mk_helpers.ps1" echo         private static readonly string PreferredTzId =
>> "__mk_helpers.ps1" echo             TryGetTimeZone("Israel Standard Time")
>> "__mk_helpers.ps1" echo             ?? TryGetTimeZone("Asia/Jerusalem")
>> "__mk_helpers.ps1" echo             ?? TimeZoneInfo.Local.Id;
>> "__mk_helpers.ps1" echo         public static void ComputeDates(GlobalVariables _context)
>> "__mk_helpers.ps1" echo         {
>> "__mk_helpers.ps1" echo             var tz = TimeZoneInfo.FindSystemTimeZoneById(PreferredTzId);
>> "__mk_helpers.ps1" echo             var now = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, tz);
>> "__mk_helpers.ps1" echo             var today = now.Date;
>> "__mk_helpers.ps1" echo             // threshold 10:01 local
>> "__mk_helpers.ps1" echo             var threshold = today.AddHours(10).AddMinutes(1);
>> "__mk_helpers.ps1" echo             var pickedDate = now ^> threshold ? today.AddDays(-1) : today;
>> "__mk_helpers.ps1" echo             _context.TodayDateYmd = pickedDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
>> "__mk_helpers.ps1" echo             var nextMonth15 = new DateTime(pickedDate.Year, pickedDate.Month, 1).AddMonths(1);
>> "__mk_helpers.ps1" echo             nextMonth15 = new DateTime(nextMonth15.Year, nextMonth15.Month, 15);
>> "__mk_helpers.ps1" echo             _context.NextMonthYmd = nextMonth15.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
>> "__mk_helpers.ps1" echo             Console.WriteLine($"todayDate = {_context.TodayDateYmd}");
>> "__mk_helpers.ps1" echo             Console.WriteLine($"nextMonth = {_context.NextMonthYmd}");
>> "__mk_helpers.ps1" echo         }
>> "__mk_helpers.ps1" echo         // =========================== HTTP initializer ===========================
>> "__mk_helpers.ps1" echo         public static class HttpEx
>> "__mk_helpers.ps1" echo         {
>> "__mk_helpers.ps1" echo             public static HttpClient Client { get; private set; } = new HttpClient();
>> "__mk_helpers.ps1" echo             public static void Initialize(HttpClient client)
>> "__mk_helpers.ps1" echo             {
>> "__mk_helpers.ps1" echo                 Client = client ?? throw new ArgumentNullException(nameof(client));
>> "__mk_helpers.ps1" echo             }
>> "__mk_helpers.ps1" echo             public static Task SleepDelay(int ms, CancellationToken cancelToken = default) =^>
>> "__mk_helpers.ps1" echo                 Task.Delay(ms, cancelToken);
>> "__mk_helpers.ps1" echo         }
>> "__mk_helpers.ps1" echo     }
>> "__mk_helpers.ps1" echo }
>> "__mk_helpers.ps1" echo '@
>> "__mk_helpers.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/Helpers.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_helpers.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create Create Helpers.cs & exit /b 1 )
del "__mk_helpers.ps1"



:: ===== Create RunAsync.cs =====
echo Creating RunAsync.cs...

> "__mk_runasync.ps1" echo $content = @'
>> "__mk_runasync.ps1" echo namespace GlobalClasses
>> "__mk_runasync.ps1" echo {
>> "__mk_runasync.ps1" echo     public static class RunAsync
>> "__mk_runasync.ps1" echo     {
>> "__mk_runasync.ps1" echo         public static async Task RunAsynchronously()
>> "__mk_runasync.ps1" echo         {
>> "__mk_runasync.ps1" echo             try
>> "__mk_runasync.ps1" echo             {
>> "__mk_runasync.ps1" echo                 await Orchestration.RunAsync();
>> "__mk_runasync.ps1" echo                 Console.WriteLine("Test procedure run is completed");
>> "__mk_runasync.ps1" echo             }
>> "__mk_runasync.ps1" echo             catch (Exception ex)
>> "__mk_runasync.ps1" echo             {
>> "__mk_runasync.ps1" echo                 Console.Error.WriteLine("Test procedure run failed: " + ex);
>> "__mk_runasync.ps1" echo                 Environment.ExitCode = 1;
>> "__mk_runasync.ps1" echo             }
>> "__mk_runasync.ps1" echo         }
>> "__mk_runasync.ps1" echo     }
>> "__mk_runasync.ps1" echo }
>> "__mk_runasync.ps1" echo '@
>> "__mk_runasync.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/RunAsync.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_runasync.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create Create RunAsync.cs & exit /b 1 )
del "__mk_runasync.ps1"



:: ===== Create Orchestration.cs =====
echo Creating Orchestration.cs...

> "__mk_orchestration.ps1" echo $content = @'
>> "__mk_orchestration.ps1" echo using System.Text.Json;
>> "__mk_orchestration.ps1" echo using Scheduling;
>> "__mk_orchestration.ps1" echo using static GlobalClasses.TimeUtil;
>> "__mk_orchestration.ps1" echo //
>> "__mk_orchestration.ps1" echo namespace GlobalClasses
>> "__mk_orchestration.ps1" echo {
>> "__mk_orchestration.ps1" echo     public class Orchestration
>> "__mk_orchestration.ps1" echo     {
>> "__mk_orchestration.ps1" echo         // No-arg wrapper uses defaults declared in the GlobalVariables
>> "__mk_orchestration.ps1" echo         public static Task RunAsync(CancellationToken cancelToken = default) =^>
>> "__mk_orchestration.ps1" echo             RunAsync(new GlobalVariables(), cancelToken);
>> "__mk_orchestration.ps1" echo         // Primary overload that uses GlobalVariables
>> "__mk_orchestration.ps1" echo         public static async Task RunAsync(GlobalVariables _context, CancellationToken cancelToken = default)
>> "__mk_orchestration.ps1" echo         {
>> "__mk_orchestration.ps1" echo             // HttpClient + JSON options
>> "__mk_orchestration.ps1" echo             var http = new HttpClient { BaseAddress = new Uri(_context.BaseUrl) };
>> "__mk_orchestration.ps1" echo             var json = new JsonSerializerOptions(JsonSerializerDefaults.Web);
>> "__mk_orchestration.ps1" echo             // Sign-in using the creds stored in GlobalVariables
>> "__mk_orchestration.ps1" echo             await _context.InitializeAsync(http, json, cancelToken).ConfigureAwait(false);
>> "__mk_orchestration.ps1" echo             // Make helpers share the same HttpClient
>> "__mk_orchestration.ps1" echo             HttpEx.Initialize(http);
>> "__mk_orchestration.ps1" echo             // === dateForControl ===
>> "__mk_orchestration.ps1" echo             TimeUtil.ComputeDates(_context);
>> "__mk_orchestration.ps1" echo             await HttpEx.SleepDelay(5000, cancelToken);
>> "__mk_orchestration.ps1" echo             // === Machine create ===
>> "__mk_orchestration.ps1" echo             await CreateMachineTask.ExecuteAsync(_context, cancelToken);
>> "__mk_orchestration.ps1" echo             await HttpEx.SleepDelay(5000, cancelToken);
>> "__mk_orchestration.ps1" echo             // === Machine update ===
>> "__mk_orchestration.ps1" echo             await UpdateMachineTask.ExecuteAsync(_context, cancelToken);
>> "__mk_orchestration.ps1" echo             await HttpEx.SleepDelay(5000, cancelToken);
>> "__mk_orchestration.ps1" echo             // === Machine tasks/delete ===
>> "__mk_orchestration.ps1" echo             // await DeleteMachineTask.ExecuteAsync(_context, cancelToken);
>> "__mk_orchestration.ps1" echo             // await HttpEx.SleepDelay(5000, cancelToken);
>> "__mk_orchestration.ps1" echo         }
>> "__mk_orchestration.ps1" echo     }
>> "__mk_orchestration.ps1" echo }
>> "__mk_orchestration.ps1" echo '@
>> "__mk_orchestration.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/Orchestration.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_orchestration.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create Orchestration.cs & exit /b 1 )
del "__mk_orchestration.ps1"



:: ===== Create GlobalVariables.cs =====
echo Creating GlobalVariables.cs...

> "__mk_global_variables.ps1" echo $content = @'
>> "__mk_global_variables.ps1" echo using System.Text.Json;
>> "__mk_global_variables.ps1" echo //
>> "__mk_global_variables.ps1" echo namespace GlobalClasses
>> "__mk_global_variables.ps1" echo {
>> "__mk_global_variables.ps1" echo     public class GlobalVariables
>> "__mk_global_variables.ps1" echo     {
>> "__mk_global_variables.ps1" echo         public string TokenKoken { get; set; } = "";
>> "__mk_global_variables.ps1" echo         public string BaseUrl { get; set; } = "https://qa-cortex.nayax.com/";
>> "__mk_global_variables.ps1" echo         public long MachineId { get; set; } = 1000757281;
>> "__mk_global_variables.ps1" echo         public string DriverId { get; set; } = "a5891545-9a28-47e6-8762-9afac1e9439c";
>> "__mk_global_variables.ps1" echo         public string TodayDateYmd { get; set; } = "";
>> "__mk_global_variables.ps1" echo         public string NextMonthYmd { get; set; } = "";
>> "__mk_global_variables.ps1" echo         public long? SchedulingId { get; set; }
>> "__mk_global_variables.ps1" echo         // If you want to keep creds here:
>> "__mk_global_variables.ps1" echo         public static string username { get; set; } = "sergeyr";
>> "__mk_global_variables.ps1" echo         public static string password { get; set; } = "ribi69qa2******";
>> "__mk_global_variables.ps1" echo         //
>> "__mk_global_variables.ps1" echo         public async Task InitializeAsync(HttpClient http, JsonSerializerOptions json, CancellationToken cancelToken = default)
>> "__mk_global_variables.ps1" echo         {
>> "__mk_global_variables.ps1" echo             // Make sure the HttpClient has BaseAddress set from your BaseUrl
>> "__mk_global_variables.ps1" echo             if (http.BaseAddress is null ^&^& BaseUrl.Length > 0)
>> "__mk_global_variables.ps1" echo                 http.BaseAddress = new Uri(BaseUrl);
>> "__mk_global_variables.ps1" echo             //
>> "__mk_global_variables.ps1" echo             var signIn = new SignIn(http, json);
>> "__mk_global_variables.ps1" echo             TokenKoken = await signIn.LogIn(username, password, cancelToken).ConfigureAwait(false);
>> "__mk_global_variables.ps1" echo         }
>> "__mk_global_variables.ps1" echo     }
>> "__mk_global_variables.ps1" echo }
>> "__mk_global_variables.ps1" echo '@
>> "__mk_global_variables.ps1" echo Set-Content -Path '%GLOBAL_CLASSES_DIR%/GlobalVariables.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_global_variables.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create GlobalVariables.cs & exit /b 1 )
del "__mk_global_variables.ps1"



:: ===== Create CreateMachineTask.cs =====
echo Creating CreateMachineTask.cs...

> "__mk_create_machine_task.ps1" echo $content = @'
>> "__mk_create_machine_task.ps1" echo using GlobalClasses;
>> "__mk_create_machine_task.ps1" echo using GlobalClasses;
>> "__mk_create_machine_task.ps1" echo using System.Net;
>> "__mk_create_machine_task.ps1" echo using System.Text;
>> "__mk_create_machine_task.ps1" echo using System.Text.Json;
>> "__mk_create_machine_task.ps1" echo using static GlobalClasses.Helpers;
>> "__mk_create_machine_task.ps1" echo //
>> "__mk_create_machine_task.ps1" echo namespace Scheduling
>> "__mk_create_machine_task.ps1" echo {
>> "__mk_create_machine_task.ps1" echo     public class CreateMachineTask
>> "__mk_create_machine_task.ps1" echo     {
>> "__mk_create_machine_task.ps1" echo         private static readonly JsonSerializerOptions JsonOpts = new(JsonSerializerDefaults.Web);
>> "__mk_create_machine_task.ps1" echo         public static async Task ExecuteAsync(GlobalVariables _context, CancellationToken cancelToken = default)
>> "__mk_create_machine_task.ps1" echo         {
>> "__mk_create_machine_task.ps1" echo             HttpEx.EnsureBearer(_context);
>> "__mk_create_machine_task.ps1" echo                 // Ensure today is computed
>> "__mk_create_machine_task.ps1" echo                 if (string.IsNullOrWhiteSpace(_context.TodayDateYmd))
>> "__mk_create_machine_task.ps1" echo                 TimeUtil.ComputeDates(_context);
>> "__mk_create_machine_task.ps1" echo             var taskLutIdArray = new[] { "887779090" /*, "887779091", "887779092"*/ };
>> "__mk_create_machine_task.ps1" echo                 foreach (var taskId in taskLutIdArray)
>> "__mk_create_machine_task.ps1" echo                 {
>> "__mk_create_machine_task.ps1" echo                     var payloadItem = BuildPayload(_context, taskId, notes: "What else beside the wealth and health?");
>> "__mk_create_machine_task.ps1" echo                     var json = JsonSerializer.Serialize(new[] { payloadItem }, JsonOpts);
>> "__mk_create_machine_task.ps1" echo                         using var req = new HttpRequestMessage(HttpMethod.Post, $"{_context.BaseUrl}/v1/scheduling/v1/schedule/machine-tasks")
>> "__mk_create_machine_task.ps1" echo                         {
>> "__mk_create_machine_task.ps1" echo                             Content = new StringContent(json, Encoding.UTF8, "application/json")
>> "__mk_create_machine_task.ps1" echo                         };
>> "__mk_create_machine_task.ps1" echo                     var res = await HttpEx.Client.SendAsync(req, cancelToken);
>> "__mk_create_machine_task.ps1" echo                     var ok = res.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created;
>> "__mk_create_machine_task.ps1" echo                     AssertEx.Require(ok, "Does POST create task response have a status OK?");
>> "__mk_create_machine_task.ps1" echo                     // Uncomment when ready
>> "__mk_create_machine_task.ps1" echo                     await GetMachineTask.ExecuteAsync(_context, sender: "POST", cancelToken);
>> "__mk_create_machine_task.ps1" echo                 }
>> "__mk_create_machine_task.ps1" echo         }
>> "__mk_create_machine_task.ps1" echo         private static TaskPayloadItem BuildPayload(GlobalVariables _context, string taskId, string notes)
>> "__mk_create_machine_task.ps1" echo         {
>> "__mk_create_machine_task.ps1" echo             return new TaskPayloadItem
>> "__mk_create_machine_task.ps1" echo             {
>> "__mk_create_machine_task.ps1" echo                 SchedulingId = null,
>> "__mk_create_machine_task.ps1" echo                 MachineId = _context.MachineId,
>> "__mk_create_machine_task.ps1" echo                 TaskLutId = taskId,
>> "__mk_create_machine_task.ps1" echo                 DriverId = _context.DriverId,
>> "__mk_create_machine_task.ps1" echo                 ScheduleDate = $"{_context.TodayDateYmd}T11:10:05.533Z",
>> "__mk_create_machine_task.ps1" echo                 StatusId = 996231363,
>> "__mk_create_machine_task.ps1" echo                 Notes = notes,
>> "__mk_create_machine_task.ps1" echo                 Duration = 88,
>> "__mk_create_machine_task.ps1" echo                 GeneratePickList = false,
>> "__mk_create_machine_task.ps1" echo                 GeneratePickListTime = null,
>> "__mk_create_machine_task.ps1" echo                 GeneratePickListRange = 159296591,
>> "__mk_create_machine_task.ps1" echo                 ScheduleNextWorkingDay = false,
>> "__mk_create_machine_task.ps1" echo                 TimezoneOffset = 0.0,
>> "__mk_create_machine_task.ps1" echo                 AssignedToPatternId = null,
>> "__mk_create_machine_task.ps1" echo                 AssignedToScheduleId = "159296591",
>> "__mk_create_machine_task.ps1" echo                 Pattern = new Pattern
>> "__mk_create_machine_task.ps1" echo                 {
>> "__mk_create_machine_task.ps1" echo                     PatternId = null,
>> "__mk_create_machine_task.ps1" echo                     Name = null,
>> "__mk_create_machine_task.ps1" echo                     Description = "string",
>> "__mk_create_machine_task.ps1" echo                     RepeatType = 1,
>> "__mk_create_machine_task.ps1" echo                     RepeatInterval = 2,
>> "__mk_create_machine_task.ps1" echo                     IsStatic = false,
>> "__mk_create_machine_task.ps1" echo                     StartOn = $"{_context.TodayDateYmd}T11:10:43.402Z",
>> "__mk_create_machine_task.ps1" echo                     EndOn = $"{_context.TodayDateYmd}T11:10:43.402Z",
>> "__mk_create_machine_task.ps1" echo                     EndAfter = 1,
>> "__mk_create_machine_task.ps1" echo                     RepeatOn = { new RepeatOn() }
>> "__mk_create_machine_task.ps1" echo                 },
>> "__mk_create_machine_task.ps1" echo                 IsMobile = true,
>> "__mk_create_machine_task.ps1" echo                 IsSeriesUpdate = false,
>> "__mk_create_machine_task.ps1" echo                 OriginalDate = null,
>> "__mk_create_machine_task.ps1" echo                 IncompletionReasonId = 0,
>> "__mk_create_machine_task.ps1" echo                 IncompletionReason = "Crazy&Lazy"
>> "__mk_create_machine_task.ps1" echo             };
>> "__mk_create_machine_task.ps1" echo         }
>> "__mk_create_machine_task.ps1" echo     }
>> "__mk_create_machine_task.ps1" echo }
>> "__mk_create_machine_task.ps1" echo '@
>> "__mk_create_machine_task.ps1" echo Set-Content -Path '%SCHEDULING_DIR%/CreateMachineTask.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_create_machine_task.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create CreateMachineTask.cs & exit /b 1 )
del "__mk_create_machine_task.ps1"



:: ===== Create GetMachineTask.cs =====
echo Creating GetMachineTask.cs...

> "__mk_get_machine_task.ps1" echo $content = @'
>> "__mk_get_machine_task.ps1" echo using System;
>> "__mk_get_machine_task.ps1" echo using System.Collections.Generic;
>> "__mk_get_machine_task.ps1" echo using System.Linq;
>> "__mk_get_machine_task.ps1" echo using System.Net;
>> "__mk_get_machine_task.ps1" echo using System.Text;
>> "__mk_get_machine_task.ps1" echo using System.Text.Json;
>> "__mk_get_machine_task.ps1" echo using System.Threading.Tasks;
>> "__mk_get_machine_task.ps1" echo using static GlobalClasses.Helpers;
>> "__mk_get_machine_task.ps1" echo //
>> "__mk_get_machine_task.ps1" echo namespace Scheduling
>> "__mk_get_machine_task.ps1" echo {
>> "__mk_get_machine_task.ps1" echo     public class GetMachineTask
>> "__mk_get_machine_task.ps1" echo     {
>> "__mk_get_machine_task.ps1" echo         private static readonly JsonSerializerOptions JsonOpts = new(JsonSerializerDefaults.Web)
>> "__mk_get_machine_task.ps1" echo         {
>> "__mk_get_machine_task.ps1" echo             PropertyNameCaseInsensitive = true
>> "__mk_get_machine_task.ps1" echo         };
>> "__mk_get_machine_task.ps1" echo         public static async Task ExecuteAsync(GlobalClasses.GlobalVariables _Context, string sender, CancellationToken ct = default)
>> "__mk_get_machine_task.ps1" echo         {
>> "__mk_get_machine_task.ps1" echo             HttpEx.EnsureBearer(_Context);
>> "__mk_get_machine_task.ps1" echo             // Assume _Context already has machineId
>> "__mk_get_machine_task.ps1" echo             var uri = $"{_Context.BaseUrl}/v1/scheduling/v1/schedule/machine-tasks" +
>> "__mk_get_machine_task.ps1" echo                       $"?MachineId={Uri.EscapeDataString(_Context.MachineId.ToString())}" +
>> "__mk_get_machine_task.ps1" echo                       $"&DriverId={Uri.EscapeDataString(_Context.DriverId)}" +
>> "__mk_get_machine_task.ps1" echo                       $"&TimePeriod=333" +
>> "__mk_get_machine_task.ps1" echo                       $"&StartDate={Uri.EscapeDataString(_Context.TodayDateYmd)}" +
>> "__mk_get_machine_task.ps1" echo                       $"&EndDate={Uri.EscapeDataString(_Context.NextMonthYmd)}" +
>> "__mk_get_machine_task.ps1" echo                       $"&IsForlist=true";
>> "__mk_get_machine_task.ps1" echo             using var req = new HttpRequestMessage(HttpMethod.Get, uri);
>> "__mk_get_machine_task.ps1" echo             var res = await HttpEx.Client.SendAsync(req, ct);
>> "__mk_get_machine_task.ps1" echo             var ok = res.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created;
>> "__mk_get_machine_task.ps1" echo             AssertEx.Require(ok, "Did GET read the task method pass with status OK?");
>> "__mk_get_machine_task.ps1" echo             var body = await res.Content.ReadAsStringAsync(ct);
>> "__mk_get_machine_task.ps1" echo //
>> "__mk_get_machine_task.ps1" echo             List^<MachineTasksEnvelope^>? envelope;
>> "__mk_get_machine_task.ps1" echo             try { envelope = JsonSerializer.Deserialize^<List^<MachineTasksEnvelope^>^>(body, JsonOpts); }
>> "__mk_get_machine_task.ps1" echo             catch { envelope = null; }
>> "__mk_get_machine_task.ps1" echo             envelope ??= new List^<MachineTasksEnvelope^>();
>> "__mk_get_machine_task.ps1" echo //
>> "__mk_get_machine_task.ps1" echo                 if (sender == "POST")
>> "__mk_get_machine_task.ps1" echo                 {
>> "__mk_get_machine_task.ps1" echo                     var firstDate = envelope.FirstOrDefault()?.ScheduleTasks?.FirstOrDefault()?.ScheduleDate ?? "";
>> "__mk_get_machine_task.ps1" echo                     AssertEx.Require(firstDate == $"{_Context.TodayDateYmd}T00:00:00",
>> "__mk_get_machine_task.ps1" echo                         "Had the task been created?");
>> "__mk_get_machine_task.ps1" echo                 }
>> "__mk_get_machine_task.ps1" echo                 else
>> "__mk_get_machine_task.ps1" echo                 {
>> "__mk_get_machine_task.ps1" echo                     var notes = envelope.FirstOrDefault()?.ScheduleTasks?.FirstOrDefault()?.Notes;
>> "__mk_get_machine_task.ps1" echo                     Console.WriteLine("Top notes: " + (notes ?? "<null>"));
>> "__mk_get_machine_task.ps1" echo                     AssertEx.Require(notes?.Contains("What?!!!") == true,
>> "__mk_get_machine_task.ps1" echo                         "Had the task been updated?");
>> "__mk_get_machine_task.ps1" echo                 }
>> "__mk_get_machine_task.ps1" echo             // Extract schedulingIds like flatMap
>> "__mk_get_machine_task.ps1" echo             var ids = envelope
>> "__mk_get_machine_task.ps1" echo                 .SelectMany(i =^> i.ScheduleTasks ?? new List^<ScheduleTask^>())
>> "__mk_get_machine_task.ps1" echo                 .Select(t =^> t.SchedulingId)
>> "__mk_get_machine_task.ps1" echo                 .ToList();
>> "__mk_get_machine_task.ps1" echo             _Context.SchedulingId = ids.FirstOrDefault();
>> "__mk_get_machine_task.ps1" echo             Console.WriteLine("schedulingIds: [" + string.Join(",", ids) + "]");
>> "__mk_get_machine_task.ps1" echo         }
>> "__mk_get_machine_task.ps1" echo     }
>> "__mk_get_machine_task.ps1" echo }
>> "__mk_get_machine_task.ps1" echo '@
>> "__mk_get_machine_task.ps1" echo Set-Content -Path '%SCHEDULING_DIR%/GetMachineTask.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_get_machine_task.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create GetMachineTask.cs & exit /b 1 )
del "__mk_get_machine_task.ps1"



:: ===== Create UpdateMachineTask.cs =====
echo Creating UpdateMachineTask.cs...

> "__mk_update_machine_task.ps1" echo $content = @'
>> "__mk_update_machine_task.ps1" echo using GlobalClasses;
>> "__mk_update_machine_task.ps1" echo using System.Net;
>> "__mk_update_machine_task.ps1" echo using System.Text;
>> "__mk_update_machine_task.ps1" echo using System.Text.Json;
>> "__mk_update_machine_task.ps1" echo using static GlobalClasses.Helpers;
>> "__mk_update_machine_task.ps1" echo //
>> "__mk_update_machine_task.ps1" echo namespace Scheduling
>> "__mk_update_machine_task.ps1" echo {
>> "__mk_update_machine_task.ps1" echo     public class UpdateMachineTask
>> "__mk_update_machine_task.ps1" echo     {
>> "__mk_update_machine_task.ps1" echo         private static readonly JsonSerializerOptions JsonOpts = new(JsonSerializerDefaults.Web);
>> "__mk_update_machine_task.ps1" echo         public static async Task ExecuteAsync(GlobalVariables _Context, CancellationToken ct = default)
>> "__mk_update_machine_task.ps1" echo         {
>> "__mk_update_machine_task.ps1" echo             HttpEx.EnsureBearer(_Context);
>> "__mk_update_machine_task.ps1" echo                 if (string.IsNullOrWhiteSpace(_Context.TodayDateYmd))
>> "__mk_update_machine_task.ps1" echo                 TimeUtil.ComputeDates(_Context);
>> "__mk_update_machine_task.ps1" echo             // Rely on SchedulingId from the GET step after POST
>> "__mk_update_machine_task.ps1" echo                 if (_Context.SchedulingId is null)
>> "__mk_update_machine_task.ps1" echo                 throw new System.InvalidOperationException("schedulingId is null. Run POST + GET first.");
>> "__mk_update_machine_task.ps1" echo             var payloadItem = new TaskPayloadItem
>> "__mk_update_machine_task.ps1" echo             {
>> "__mk_update_machine_task.ps1" echo                 SchedulingId = _Context.SchedulingId,
>> "__mk_update_machine_task.ps1" echo                 MachineId = _Context.MachineId,
>> "__mk_update_machine_task.ps1" echo                 TaskLutId = "887779090",
>> "__mk_update_machine_task.ps1" echo                 DriverId = _Context.DriverId,
>> "__mk_update_machine_task.ps1" echo                 ScheduleDate = $"{_Context.TodayDateYmd}T11:10:05.533Z",
>> "__mk_update_machine_task.ps1" echo                 StatusId = 996231363,
>> "__mk_update_machine_task.ps1" echo                 Notes = "What?!!!",
>> "__mk_update_machine_task.ps1" echo                 Duration = 88,
>> "__mk_update_machine_task.ps1" echo                 GeneratePickList = false,
>> "__mk_update_machine_task.ps1" echo                 GeneratePickListTime = null,
>> "__mk_update_machine_task.ps1" echo                 GeneratePickListRange = 159296591,
>> "__mk_update_machine_task.ps1" echo                 ScheduleNextWorkingDay = false,
>> "__mk_update_machine_task.ps1" echo                 TimezoneOffset = 0.0,
>> "__mk_update_machine_task.ps1" echo                 AssignedToPatternId = null,
>> "__mk_update_machine_task.ps1" echo                 AssignedToScheduleId = "159296591",
>> "__mk_update_machine_task.ps1" echo                 Pattern = new Pattern
>> "__mk_update_machine_task.ps1" echo                 {
>> "__mk_update_machine_task.ps1" echo                     PatternId = null,
>> "__mk_update_machine_task.ps1" echo                     Name = null,
>> "__mk_update_machine_task.ps1" echo                     Description = "string",
>> "__mk_update_machine_task.ps1" echo                     RepeatType = 1,
>> "__mk_update_machine_task.ps1" echo                     RepeatInterval = 2,
>> "__mk_update_machine_task.ps1" echo                     IsStatic = false,
>> "__mk_update_machine_task.ps1" echo                     StartOn = $"{_Context.TodayDateYmd}T11:10:43.402Z",
>> "__mk_update_machine_task.ps1" echo                     EndOn = $"{_Context.TodayDateYmd}T11:10:43.402Z",
>> "__mk_update_machine_task.ps1" echo                     EndAfter = 1,
>> "__mk_update_machine_task.ps1" echo                     RepeatOn = { new RepeatOn() }
>> "__mk_update_machine_task.ps1" echo                 },
>> "__mk_update_machine_task.ps1" echo                 IsMobile = true,
>> "__mk_update_machine_task.ps1" echo                 IsSeriesUpdate = false,
>> "__mk_update_machine_task.ps1" echo                 OriginalDate = null,
>> "__mk_update_machine_task.ps1" echo                 IncompletionReasonId = 0,
>> "__mk_update_machine_task.ps1" echo                 IncompletionReason = "Crazy&Lazy"
>> "__mk_update_machine_task.ps1" echo             };
>> "__mk_update_machine_task.ps1" echo             var json = JsonSerializer.Serialize(new[] { payloadItem }, JsonOpts);
>> "__mk_update_machine_task.ps1" echo             using var req = new HttpRequestMessage(HttpMethod.Put, $"{_Context.BaseUrl}/v1/scheduling/v1/schedule/machine-tasks")
>> "__mk_update_machine_task.ps1" echo             {
>> "__mk_update_machine_task.ps1" echo                 Content = new StringContent(json, Encoding.UTF8, "application/json")
>> "__mk_update_machine_task.ps1" echo             };
>> "__mk_update_machine_task.ps1" echo             var res = await HttpEx.Client.SendAsync(req, ct);
>> "__mk_update_machine_task.ps1" echo             var ok = res.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created;
>> "__mk_update_machine_task.ps1" echo             AssertEx.Require(ok, "Does PUT update task response have a status OK?");
>> "__mk_update_machine_task.ps1" echo             //Uncomment when ready
>> "__mk_update_machine_task.ps1" echo             await GetMachineTask.ExecuteAsync(_Context, sender: "PUT", ct);
>> "__mk_update_machine_task.ps1" echo         }
>> "__mk_update_machine_task.ps1" echo     }
>> "__mk_update_machine_task.ps1" echo }
>> "__mk_update_machine_task.ps1" echo '@
>> "__mk_update_machine_task.ps1" echo Set-Content -Path '%SCHEDULING_DIR%/UpdateMachineTask.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_update_machine_task.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create UpdateMachineTask.cs & exit /b 1 )
del "__mk_update_machine_task.ps1"



:: ===== Create DeleteMachineTask.cs =====
echo Creating DeleteMachineTask.cs...

> "__mk_delete_machine_task.ps1" echo $content = @'
>> "__mk_delete_machine_task.ps1" echo using System.Net;
>> "__mk_delete_machine_task.ps1" echo using System.Text.Json;
>> "__mk_delete_machine_task.ps1" echo using static GlobalClasses.Helpers;
>> "__mk_delete_machine_task.ps1" echo //
>> "__mk_delete_machine_task.ps1" echo namespace Scheduling
>> "__mk_delete_machine_task.ps1" echo {
>> "__mk_delete_machine_task.ps1" echo     public class DeleteMachineTask
>> "__mk_delete_machine_task.ps1" echo     {
>> "__mk_delete_machine_task.ps1" echo         private static readonly JsonSerializerOptions JsonOpts = new(JsonSerializerDefaults.Web)
>> "__mk_delete_machine_task.ps1" echo         {
>> "__mk_delete_machine_task.ps1" echo             PropertyNameCaseInsensitive = true
>> "__mk_delete_machine_task.ps1" echo         };
>> "__mk_delete_machine_task.ps1" echo         public static async Task ExecuteAsync(GlobalClasses.GlobalVariables _Context, string sender, CancellationToken ct = default)
>> "__mk_delete_machine_task.ps1" echo         {
>> "__mk_delete_machine_task.ps1" echo             HttpEx.EnsureBearer(_Context);
>> "__mk_delete_machine_task.ps1" echo             // Assume _Context already has machineId
>> "__mk_delete_machine_task.ps1" echo             var uri = $"{_Context.BaseUrl}/v1/scheduling/v1/schedule/machine-tasks" +
>> "__mk_delete_machine_task.ps1" echo                       $"?MachineId={Uri.EscapeDataString(_Context.MachineId.ToString())}" +
>> "__mk_delete_machine_task.ps1" echo                       $"&DriverId={Uri.EscapeDataString(_Context.DriverId)}" +
>> "__mk_delete_machine_task.ps1" echo                       $"&TimePeriod=333" +
>> "__mk_delete_machine_task.ps1" echo                       $"&StartDate={Uri.EscapeDataString(_Context.TodayDateYmd)}" +
>> "__mk_delete_machine_task.ps1" echo                       $"&EndDate={Uri.EscapeDataString(_Context.NextMonthYmd)}" +
>> "__mk_delete_machine_task.ps1" echo                       $"&IsForlist=true";
>> "__mk_delete_machine_task.ps1" echo             //
>> "__mk_delete_machine_task.ps1" echo             using var req = new HttpRequestMessage(HttpMethod.Get, uri);
>> "__mk_delete_machine_task.ps1" echo             var res = await HttpEx.Client.SendAsync(req, ct);
>> "__mk_delete_machine_task.ps1" echo             var ok = res.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created;
>> "__mk_delete_machine_task.ps1" echo             AssertEx.Require(ok, "Did GET read the task method pass with status OK?");
>> "__mk_delete_machine_task.ps1" echo             var body = await res.Content.ReadAsStringAsync(ct);
>> "__mk_delete_machine_task.ps1" echo             List^<MachineTasksEnvelope^>? envelope;
>> "__mk_delete_machine_task.ps1" echo             try { envelope = JsonSerializer.Deserialize^<List^<MachineTasksEnvelope^>^>(body, JsonOpts); }
>> "__mk_delete_machine_task.ps1" echo             catch { envelope = null; }
>> "__mk_delete_machine_task.ps1" echo             envelope ??= new List^<MachineTasksEnvelope^>();
>> "__mk_delete_machine_task.ps1" echo                 if (sender == "POST")
>> "__mk_delete_machine_task.ps1" echo                 {
>> "__mk_delete_machine_task.ps1" echo                     var firstDate = envelope.FirstOrDefault()?.ScheduleTasks?.FirstOrDefault()?.ScheduleDate ?? "";
>> "__mk_delete_machine_task.ps1" echo                     AssertEx.Require(firstDate == $"{_Context.TodayDateYmd}T00:00:00",
>> "__mk_delete_machine_task.ps1" echo                         "Had the task been created?");
>> "__mk_delete_machine_task.ps1" echo                 }
>> "__mk_delete_machine_task.ps1" echo                 else
>> "__mk_delete_machine_task.ps1" echo                 {
>> "__mk_delete_machine_task.ps1" echo                     var notes = envelope.FirstOrDefault()?.ScheduleTasks?.FirstOrDefault()?.Notes;
>> "__mk_delete_machine_task.ps1" echo                     Console.WriteLine("Top notes: " + (notes ?? "<null>"));
>> "__mk_delete_machine_task.ps1" echo                     AssertEx.Require(notes?.Contains("What?!!!") == true,
>> "__mk_delete_machine_task.ps1" echo                         "Had the task been updated?");
>> "__mk_delete_machine_task.ps1" echo                 }
>> "__mk_delete_machine_task.ps1" echo             // Extract schedulingIds like flatMap in JS
>> "__mk_delete_machine_task.ps1" echo             var ids = envelope
>> "__mk_delete_machine_task.ps1" echo                 .SelectMany(i =^> i.ScheduleTasks ?? new List^<ScheduleTask^>())
>> "__mk_delete_machine_task.ps1" echo                 .Select(t =^> t.SchedulingId)
>> "__mk_delete_machine_task.ps1" echo                 .ToList();
>> "__mk_delete_machine_task.ps1" echo             _Context.SchedulingId = ids.FirstOrDefault();
>> "__mk_delete_machine_task.ps1" echo             Console.WriteLine("schedulingIds: [" + string.Join(",", ids) + "]");
>> "__mk_delete_machine_task.ps1" echo         }
>> "__mk_delete_machine_task.ps1" echo     }
>> "__mk_delete_machine_task.ps1" echo }
>> "__mk_delete_machine_task.ps1" echo '@
>> "__mk_delete_machine_task.ps1" echo Set-Content -Path '%SCHEDULING_DIR%/DeleteMachineTask.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_delete_machine_task.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create DeleteMachineTask.cs & exit /b 1 )
del "__mk_delete_machine_task.ps1"



:: ===== Create ScheduleModels.cs =====
echo Creating ScheduleModels.cs...

> "__mk_schedule_models.ps1" echo $content = @'
>> "__mk_schedule_models.ps1" echo using System.Text.Json.Serialization;
>> "__mk_schedule_models.ps1" echo //
>> "__mk_schedule_models.ps1" echo namespace Scheduling
>> "__mk_schedule_models.ps1" echo {
>> "__mk_schedule_models.ps1" echo     // One model for both Create/Update; Update will just fill SchedulingId.
>> "__mk_schedule_models.ps1" echo     public sealed class TaskPayloadItem
>> "__mk_schedule_models.ps1" echo     {
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("schedulingId")] public long? SchedulingId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("machineId")] public long MachineId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("taskLutId")] public string TaskLutId { get; set; } = "";
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("driverId")] public string DriverId { get; set; } = "";
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("scheduleDate")] public string ScheduleDate { get; set; } = "";
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("statusId")] public long StatusId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("notes")] public string? Notes { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("duration")] public int Duration { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("generatePickList")] public bool GeneratePickList { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("generatePickListTime")] public string? GeneratePickListTime { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("generatePickListRange")] public long? GeneratePickListRange { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("scheduleNextWorkingDay")] public bool ScheduleNextWorkingDay { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("timezoneOffset")] public double TimezoneOffset { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("assignedToPatternId")] public long? AssignedToPatternId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("assignedToScheduleId")] public string? AssignedToScheduleId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("pattern")] public Pattern? Pattern { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("isMobile")] public bool IsMobile { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("isSeriesUpdate")] public bool IsSeriesUpdate { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("originalDate")] public string? OriginalDate { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("incompletionReasonId")] public int IncompletionReasonId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("incompletionReason")] public string? IncompletionReason { get; set; }
>> "__mk_schedule_models.ps1" echo     }
>> "__mk_schedule_models.ps1" echo     public sealed class Pattern
>> "__mk_schedule_models.ps1" echo     {
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("patternId")] public long? PatternId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("name")] public string? Name { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("description")] public string? Description { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("repeatType")] public int RepeatType { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("repeatInterval")] public int RepeatInterval { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("isStatic")] public bool IsStatic { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("startOn")] public string? StartOn { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("endOn")] public string? EndOn { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("endAfter")] public int EndAfter { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("repeatOn")] public List^<RepeatOn^> RepeatOn { get; set; } = new();
>> "__mk_schedule_models.ps1" echo     }
>> "__mk_schedule_models.ps1" echo     public sealed class RepeatOn
>> "__mk_schedule_models.ps1" echo     {
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("tasks")] public object? Tasks { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("days")] public object? Days { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("week")] public object? Week { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("monthDay")] public object? MonthDay { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("setPos")] public object? SetPos { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("recurrenceEx")] public object? RecurrenceEx { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("weekNumber")] public object? WeekNumber { get; set; }
>> "__mk_schedule_models.ps1" echo     }
>> "__mk_schedule_models.ps1" echo // =========================== Models ===========================
>> "__mk_schedule_models.ps1" echo     public sealed class MachineTasksEnvelope
>> "__mk_schedule_models.ps1" echo     {
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("machineId")] public long MachineId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("scheduleTasks")] public List^<ScheduleTask^> ScheduleTasks { get; set; } = new();
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("schedulingId")] public long SchedulingId { get; set; }
>> "__mk_schedule_models.ps1" echo     }
>> "__mk_schedule_models.ps1" echo     public sealed class ScheduleTask
>> "__mk_schedule_models.ps1" echo     {
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("schedulingId")] public long SchedulingId { get; set; }
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("scheduleDate")] public string ScheduleDate { get; set; } = "";
>> "__mk_schedule_models.ps1" echo         [JsonPropertyName("notes")] public string? Notes { get; set; }
>> "__mk_schedule_models.ps1" echo     }
>> "__mk_schedule_models.ps1" echo }
>> "__mk_schedule_models.ps1" echo '@
>> "__mk_schedule_models.ps1" echo Set-Content -Path '%SCHEDULING_DIR%/ScheduleModels.cs' -Value $content -Encoding UTF8

powershell -NoProfile -ExecutionPolicy Bypass -File "__mk_schedule_models.ps1"
if errorlevel 1 ( echo [ERROR] Failed to create ScheduleModels.cs & exit /b 1 )
del "__mk_schedule_models.ps1"



:: ===== Ready files counter =====
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
set /p OPEN_VS=Open the project in Visual Studio? [Y/N]: 
if "%OPEN_VS%"=="y" (
    echo Opening project: %CSPROJ%
    start "" "%CSPROJ%"
) else (
    echo Project creation complete. Visual Studio not opened.
)



pause
endlocal
