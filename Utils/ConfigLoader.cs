using Microsoft.Extensions.Configuration;

namespace RemoteWinAppAutomation.Utils
{
    public static class ConfigLoader
    {
        public static IConfigurationRoot LoadConfig()
        {
            var builder = new ConfigurationBuilder()
                .AddJsonFile(Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json"), optional: false, reloadOnChange: true);
            return builder.Build();
        }
    }
}
