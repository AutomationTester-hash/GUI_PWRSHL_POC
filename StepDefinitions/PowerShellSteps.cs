
        
        
using System;
using TechTalk.SpecFlow;
using Microsoft.Extensions.Configuration;
using RemoteWinAppAutomation.Utils;
using RemoteWinAppAutomation.Pages;
using AventStack.ExtentReports;
using OpenQA.Selenium.Appium.Windows;
// using NUnit.Framework; // Removed for MsTest/SpecFlow only

namespace RemoteWinAppAutomation.StepDefinitions
{
    [Binding]
    public class PowerShellSteps
    {
        private static System.Diagnostics.Process _virtualBoxProcess;
        [Given("I open Oracle VirtualBox Manager")]
    // Opens the Oracle VirtualBox Manager application.
        public void GivenIOpenOracleVirtualBoxManager()
        {
            Console.WriteLine("[PowerShellSteps] Opening Oracle VirtualBox Manager");
            // You may want to get this from config, for now hardcode or use your config helper
            var vboxPath = _config?["PowerShell:VBoxManagePath"];
            if (string.IsNullOrEmpty(vboxPath))
                vboxPath = "C:\\Program Files\\Oracle\\VirtualBox\\VBoxManage.exe";
            var vboxUIPath = vboxPath.Replace("VBoxManage.exe", "VirtualBox.exe");
            _virtualBoxProcess = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = vboxUIPath,
                    UseShellExecute = true
                }
            };
            _virtualBoxProcess.Start();
            System.Threading.Thread.Sleep(3000);
            if (_virtualBoxProcess.HasExited)
                throw new Exception($"VirtualBox Manager did not start. Tried path: {vboxUIPath}");
        }

        [When("I close Oracle VirtualBox Manager")]
    // Closes the Oracle VirtualBox Manager application if it is running.
        public void WhenICloseOracleVirtualBoxManager()
        {
            if (_virtualBoxProcess != null && !_virtualBoxProcess.HasExited)
            {
                Console.WriteLine($"[PowerShellSteps] Attempting to close VirtualBox Manager (PID: {_virtualBoxProcess.Id})");
                _virtualBoxProcess.Kill();
                _virtualBoxProcess.WaitForExit(5000);
                Console.WriteLine("[PowerShellSteps] Closed VirtualBox Manager via process reference.");
            }
            else
            {
                Console.WriteLine("[PowerShellSteps] VirtualBox Manager process was not running or already closed.");
            }
    }
        [When(@"I wait for 5 seconds after opening custom app")]
    // Waits for 5 seconds, then tries to close the VirtualBox Manager application.
        public void WhenIWaitFor5SecondsAfterOpeningCustomApp()
        {
            var startTime = DateTime.Now;
            Console.WriteLine($"[PowerShellSteps] Starting 5 second wait before closing VirtualBox Manager at {startTime:HH:mm:ss.fff}");
            System.Threading.Thread.Sleep(5000);
            var procs = System.Diagnostics.Process.GetProcessesByName("VirtualBox");
            Console.WriteLine($"[PowerShellSteps] Found {procs.Length} VirtualBox.exe process(es) to close.");
            foreach (var proc in procs)
            {
                try
                {
                    Console.WriteLine($"[PowerShellSteps] Attempting to kill process Id={proc.Id}, Name={proc.ProcessName}");
                    try { proc.Kill(true); } catch { proc.Kill(); } // Try forceful kill if available
                    // Wait up to 5 seconds for process to exit
                    bool exited = proc.WaitForExit(5000);
                    if (exited)
                        Console.WriteLine($"[PowerShellSteps] Successfully killed process Id={proc.Id}");
                    else
                        Console.WriteLine($"[PowerShellSteps] Process Id={proc.Id} did not exit after kill attempt");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[PowerShellSteps] Failed to kill process Id={proc.Id}: {ex.Message}");
                }
            }
            var endTime = DateTime.Now;
            Console.WriteLine($"[PowerShellSteps] Finished VirtualBox Manager close step at {endTime:HH:mm:ss.fff}, duration: {(endTime-startTime).TotalSeconds:F2} seconds");
        }
        [When(@"I run the executable from config key '(.*)'")]
    // Runs a program specified in the configuration file.
        public void WhenIRunTheExecutableFromConfigKey(string configKey)
        {
            string exePath = _config[configKey];
            if (string.IsNullOrEmpty(exePath))
            {
                // Try as top-level key if not found as nested
                var keyParts = configKey.Split(':');
                if (keyParts.Length == 2)
                {
                    exePath = _config[keyParts[1]];
                }
            }
            if (string.IsNullOrEmpty(exePath))
            {
                throw new Exception($"Config key '{configKey}' not found or empty in appsettings.json");
            }
            // Reuse the existing logic for launching executables
            WhenIRunTheExecutable(exePath);
        }
        [When(@"I wait for 5 seconds after opening Calculator")]
    // Waits for 5 seconds after opening the Calculator application.
        public void WhenIWaitFor5SecondsAfterOpeningCalculator()
        {
            System.Threading.Thread.Sleep(5000);
        }
        [AfterScenario]
    // Cleans up and closes any open applications after each test scenario.
        public void AfterScenarioCleanup()
        {
            // Attempt to close Notepad, Calculator, VirtualBox, and CustomApp if running
            // Add UWP Calculator process names as well
            string[] processNames = { "notepad", "calc", "CalculatorApp", "ApplicationFrameHost", "VirtualBox" };
            foreach (var name in processNames)
            {
                var procs = System.Diagnostics.Process.GetProcessesByName(name);
                foreach (var proc in procs) { try { proc.Kill(); } catch { } }
            }

            var customAppPath = _config?["PowerShell:CustomAppPath"];
            if (!string.IsNullOrEmpty(customAppPath))
            {
                var exeName = System.IO.Path.GetFileNameWithoutExtension(customAppPath);
                var customProcs = System.Diagnostics.Process.GetProcessesByName(exeName);
                foreach (var proc in customProcs) { try { proc.Kill(); } catch { } }
            }
        }
        private static IConfiguration _config;
        private static ExtentReports _report;
        private static WindowsDriver<WindowsElement> _driver;
        private static PowerShellPage _psPage;
        private static ExtentTest _test;
        private string _output;

        [BeforeTestRun]
    // Sets up configuration and reporting before any tests run.
        public static void BeforeTestRun()
        {
            Console.WriteLine("[PowerShellSteps] BeforeTestRun executing - test framework is running.");
            _config = ConfigLoader.LoadConfig();
            _report = ReportManager.GetReporter(_config);
        }

    [Given(@"I have initialized the application for automation")]
    // Initializes the application and test session for automation.
    public void GivenIHaveInitializedTheApplicationForAutomation()
        {
            Console.WriteLine("[PowerShellSteps] GivenIHaveConnectedToTheRemoteVMPowerShell executing.");
            Console.WriteLine("[PowerShellSteps] Connecting to PowerShell on remote/local machine...");
            try
            {
                _driver = DriverFactory.CreateSession(_config);
                if (_driver != null)
                {
                    Console.WriteLine("[PowerShellSteps] Driver session established.");
                    try
                    {
                        var windowHandle = _driver.CurrentWindowHandle;
                        Console.WriteLine($"[PowerShellSteps] Current window handle: {windowHandle}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[PowerShellSteps] Could not get window handle: {ex.Message}");
                        Console.WriteLine($"[PowerShellSteps] StackTrace: {ex.StackTrace}");
                    }
                }
                else
                {
                    Console.WriteLine("[PowerShellSteps] Driver session is null!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PowerShellSteps] Exception during session creation: {ex.Message}");
                Console.WriteLine($"[PowerShellSteps] StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"[PowerShellSteps] InnerException: {ex.InnerException.Message}");
                    Console.WriteLine($"[PowerShellSteps] InnerException StackTrace: {ex.InnerException.StackTrace}");
                }
            }
            _psPage = new PowerShellPage(_driver);
            _test = _report.CreateTest("PowerShell Automation");
        }

    [When(@"I enter the command ""(.*)""")]
        // Enters a command into the PowerShell window.
        public void WhenIEnterTheCommand(string command)
        {
            if (_psPage == null)
            {
                throw new InvalidOperationException("PowerShellPage is not initialized. Session creation may have failed.");
            }
            if (_driver == null)
            {
                throw new InvalidOperationException("Driver is not initialized. Session creation may have failed.");
            }
            _psPage.EnterCommand(command);
            // Optionally, wait for output
        }
    [Given(@"I set the PowerShell execution policy to Unrestricted")]
        // Sets the PowerShell execution policy to allow all scripts to run.
        public void GivenISetThePowerShellExecutionPolicyToUnrestricted()
        {
            RemoteWinAppAutomation.Utils.PowerShellRunner.RunCommand("Set-ExecutionPolicy Unrestricted -Scope Process -Force");
        }

    [Then(@"the output should contain ""(.*)""")]
        // Checks if the output contains the expected text.
        public void ThenTheOutputShouldContain(string expected)
        {
            _output = _psPage.GetOutput();
            if (!_output.Contains(expected))
            {
                throw new Exception($"Expected output to contain: {expected}");
            }
            _test.Pass($"Output contains expected text: {expected}");

        // --- PowerShell direct execution steps ---
        }

         [When(@"I erase the text in Notepad after 2 seconds")]
    // Waits for 2 seconds, then erases all text in Notepad.
        public void WhenIEraseTheTextInNotepadAfter2Seconds()
        {
            System.Threading.Thread.Sleep(2000);
            if (_driver == null)
                throw new Exception("Driver is not initialized.");
            var edit = _driver.FindElementByClassName("Edit");
            edit.SendKeys(OpenQA.Selenium.Keys.Control + "a");
            System.Threading.Thread.Sleep(500);
            edit.SendKeys(OpenQA.Selenium.Keys.Delete);
        }

        [When(@"I run the executable '(.*)'")]
    // Runs the specified executable file.
        public void WhenIRunTheExecutable(string exePath)
        {
            // Run the .exe and capture output
            if (exePath.EndsWith("calc.exe", StringComparison.OrdinalIgnoreCase))
            {
                // Try launching calc.exe
                _output = RemoteWinAppAutomation.Utils.PowerShellRunner.RunCommand($"Start-Process '{exePath}'");
                // Wait a moment to see if Calculator process appears
                System.Threading.Thread.Sleep(1500);
                var calcProcs = System.Diagnostics.Process.GetProcessesByName("CalculatorApp");
                if (calcProcs.Length == 0)
                {
                    // Fallback: try UWP method
                    _output = RemoteWinAppAutomation.Utils.PowerShellRunner.RunCommand("start calculator:");
                }
            }
            else
            {
                _output = RemoteWinAppAutomation.Utils.PowerShellRunner.RunCommand($"& '{exePath}'");
            }
        }

        [Then(@"the executable output should contain '(.*)'")]
    // Checks if the output from the executable contains the expected text.
        public void ThenTheExecutableOutputShouldContain(string expected)
        {
            Console.WriteLine($"[PowerShellSteps] Executable output: {_output}");
            if (_test != null)
            {
                _test.Info($"Executable output: {_output}");
            }
            string normalizedOutput = string.Concat(_output.Where(c => !char.IsWhiteSpace(c)));
            string normalizedExpected = string.Concat(expected.Where(c => !char.IsWhiteSpace(c)));
            if (!normalizedOutput.Contains(normalizedExpected))
            {
                throw new Exception($"Expected executable output to contain: {expected}\nActual: {_output}");
            }
        }

        [When(@"I run the PowerShell command '(.*)'")]
    // Runs the specified PowerShell command.
        public void WhenIRunThePowerShellCommand(string command)
        {
            _output = RemoteWinAppAutomation.Utils.PowerShellRunner.RunCommand(command);
        }

        [Then(@"the PowerShell output should contain '(.*)'")]
    // Checks if the PowerShell command output contains the expected text.
        public void ThenThePowerShellOutputShouldContain(string expected)
        {
            Console.WriteLine($"[PowerShellSteps] PowerShell output: {_output}");
            string normalizedOutput = string.Concat(_output.Where(c => !char.IsWhiteSpace(c)));
            string normalizedExpected = string.Concat(expected.Where(c => !char.IsWhiteSpace(c)));
            if (normalizedOutput.Contains(normalizedExpected))
            {
                Console.WriteLine($"✅ Test Passed: Output contains '{expected}'.");
                if (_test != null)
                {
                    _test.Pass($"Output contains expected text: {expected}");
                }
            }
            else
            {
                throw new Exception($"Expected PowerShell output to contain: {expected}\nActual: {_output}");
            }
        }

        [AfterTestRun]
    // Cleans up the test session and saves the report after all tests are done.
        public static void AfterTestRun()
        {
            _driver?.Quit();
            _report?.Flush();
        }
    }
}
