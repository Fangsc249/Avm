
using System;
using Microsoft.Extensions.Configuration;

namespace AutomationVM.Core
{
    public class AppConfig
    {
        private static readonly Lazy<IConfigurationRoot> _lazyConfig = new Lazy<IConfigurationRoot>(() =>
            new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build()
        );

        private static IConfigurationRoot configRoot => _lazyConfig.Value;

        public static string ConnectionString
        {
            get
            {
                var conn = configRoot.GetValue<string>($"{Site}:ConnectionString");
                if (!conn.Contains("Data"))
                {
                    conn = Encryptor.Decrypt(conn);
                }
                return conn;
            }
        }
        public static string Site => configRoot.GetValue<string>("Site");
        public static bool WinForm => configRoot.GetValue<bool>($"{Site}:WinForm");
        public static int LogLevel => configRoot.GetValue<int>($"{Site}:LogLevel");

        public static string OutputDir(string userName)
        {
            var outDir = configRoot.GetValue<string>($"{Site}:OutputDirs:{userName}");
            if (outDir != null)
            {
                return outDir;
            }

            return configRoot.GetValue<string>($"{Site}:OutputDirs:Default");
        }
        public static string PrinterName(string outputDir)
        {
            var printer = configRoot.GetValue<string>($"{Site}:Printers:{outputDir}");
            if (printer != null)
            {
                return printer;
            }
            return configRoot.GetValue<string>($"{Site}:Printers:Default");
        }
    }
    
}
