using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Fclp;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace PECmd
{
    class Program
    {
        private static Logger _logger;

        private static bool CheckForDotnet46()
        {
            using (
                var ndpKey =
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                        .OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\\"))
            {
                var releaseKey = Convert.ToInt32(ndpKey.GetValue("Release"));

                return releaseKey >= 393295;
            }
        }

        static void Main(string[] args)
        {
            SetupNLog();


            _logger = LogManager.GetCurrentClassLogger();

            if (!CheckForDotnet46())
            {
                _logger.Warn(".net 4.6 not detected. Please install .net 4.6 and try again.");
                return;
            }

            var p = new FluentCommandLineParser<ApplicationArguments>
            {
                IsCaseSensitive = false
            };

            _logger.Warn("test");

            Console.ReadKey();
        }

        private static void SetupNLog()
        {
            var config = new LoggingConfiguration();
            var loglevel = LogLevel.Info;

            var layout = @"${message}";

            var consoleTarget = new ColoredConsoleTarget();

            config.AddTarget("console", consoleTarget);

            consoleTarget.Layout = layout;

            var rule1 = new LoggingRule("*", loglevel, consoleTarget);
            config.LoggingRules.Add(rule1);

            LogManager.Configuration = config;
        }
    }

    internal class ApplicationArguments
    {
//        public string File { get; set; }
//        public string Directory { get; set; }
//        public string SaveTo { get; set; } = string.Empty;
//        public bool GetAscii { get; set; } = true;
//        public bool GetUnicode { get; set; } = true;
//        public string LookForString { get; set; } = string.Empty;
//        public string LookForRegex { get; set; } = string.Empty;
//        public int MinimumLength { get; set; } = 3;
//        public int MaximumLength { get; set; } = -1;
//        public int BlockSizeMB { get; set; } = 512;
//        public bool ShowOffset { get; set; } = false;
//        public bool SortLength { get; set; } = false;
//        public bool SortAlpha { get; set; } = false;
//        public bool Quiet { get; set; } = false;
//        public bool GetPatterns { get; set; } = false;
    }
}
