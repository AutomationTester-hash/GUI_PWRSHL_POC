using System.Diagnostics;

namespace RemoteWinAppAutomation.Utils
{
    public static class PowerShellRunner
    {
        public static string RunCommand(string command)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-Command \"{command}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                if (!string.IsNullOrEmpty(error))
                {
                    return $"ERROR: {error}\nOUTPUT: {output}";
                }
                return output;
            }
        }
    }
}
