using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Microsoft.Extensions.Configuration;

namespace RemoteWinAppAutomation.Utils
{
    public static class DriverFactory
    {
        public static WindowsDriver<WindowsElement> CreateSession(IConfiguration config)
        {
            var appiumOptions = new AppiumOptions();
            appiumOptions.AddAdditionalCapability("app", config["PowerShell:AppPath"]);
            appiumOptions.AddAdditionalCapability("platformName", "Windows");
            appiumOptions.AddAdditionalCapability("deviceName", "WindowsPC");
            // Try to ensure window is not minimized (WinAppDriver may not support this directly)
            // appiumOptions.AddAdditionalCapability("appArguments", "-NoExit");

            string remoteUrl = $"http://{config["VM:IPAddress"]}:{config["VM:WinAppDriverPort"]}";
            Console.WriteLine($"[DriverFactory] Attempting to create session at {remoteUrl} for app {config["PowerShell:AppPath"]}");
            try
            {
                var driver = new WindowsDriver<WindowsElement>(new Uri(remoteUrl), appiumOptions);
                Console.WriteLine("[DriverFactory] Session created successfully.");
                return driver;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DriverFactory] Failed to create session: {ex.Message}");
                Console.WriteLine($"[DriverFactory] StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"[DriverFactory] InnerException: {ex.InnerException.Message}");
                    Console.WriteLine($"[DriverFactory] InnerException StackTrace: {ex.InnerException.StackTrace}");
                }
                throw;
            }
        }
    }
}
