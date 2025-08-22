using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Microsoft.Extensions.Configuration;

namespace RemoteWinAppAutomation.Utils
{
    public static class ReportManager
    {
        private static ExtentReports _extent;
        public static ExtentReports GetReporter(IConfiguration config)
        {
            if (_extent == null)
            {
                var dir = config["Report:OutputDirectory"] ?? "Reports";
                Directory.CreateDirectory(dir);
                var htmlReporter = new AventStack.ExtentReports.Reporter.ExtentSparkReporter(Path.Combine(dir, "TestReport.html"));
                _extent = new ExtentReports();
                _extent.AttachReporter(htmlReporter);
            }
            return _extent;
        }
    }
}
