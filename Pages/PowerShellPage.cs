using OpenQA.Selenium.Appium.Windows;

namespace RemoteWinAppAutomation.Pages
{
    public class PowerShellPage
    {
        private readonly WindowsDriver<WindowsElement> _driver;

        public PowerShellPage(WindowsDriver<WindowsElement> driver)
        {
            _driver = driver;
        }

        public void EnterCommand(string command)
        {
            var input = _driver.SwitchTo().ActiveElement();
            input.SendKeys(command + "\n");
        }

        public string GetOutput()
        {
            // For Notepad, get the text area and return its value
            try
            {
                var edit = _driver.FindElementByClassName("Edit");
                return edit.Text;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
