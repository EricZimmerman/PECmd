using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Fclp;
using Fclp.Internals.Extensions;
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

            p.Setup(arg => arg.File)
              .As('f')
              .WithDescription("File to search. Either this or -d is required");

            p.Setup(arg => arg.Directory)
                .As('d')
                .WithDescription("Directory to recursively process. Either this or -f is required");

            var header =
             $"PECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
             "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
             "\r\nhttps://github.com/EricZimmerman/PECmd";

            var footer = @"Examples: PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf""" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someFile.txt"" --lr guid" + "\r\n\t " +
//                         @" PECmd.exe -d ""C:\Temp"" --ls test" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa -m 15 -x 22" + "\r\n\t " +
                         @" PECmd.exe -d ""C:\Temp""" + "\r\n\t ";

            p.SetupHelp("?", "help").WithHeader(header).Callback(text => _logger.Info(text + "\r\n" + footer));

            var result = p.Parse(args);

            if (result.HelpCalled)
            {
                return;
            }

            if (result.HasErrors)
            {
                _logger.Error("");
                _logger.Error(result.ErrorText);

                p.HelpOption.ShowHelp(p.Options);

                return;
            }

            if (p.Object.File.IsNullOrEmpty() && p.Object.Directory.IsNullOrEmpty())
            {
                p.HelpOption.ShowHelp(p.Options);

                _logger.Warn("Either -f or -d is required. Exiting");
                return;
            }

            if (p.Object.File.IsNullOrEmpty() == false && !File.Exists(p.Object.File))
            {
                _logger.Warn($"File '{p.Object.File}' not found. Exiting");
                return;
            }

            if (p.Object.Directory.IsNullOrEmpty() == false && !Directory.Exists(p.Object.Directory))
            {
                _logger.Warn($"Directory '{p.Object.Directory}' not found. Exiting");
                return;
            }

            if (p.Object.File?.Length > 0)
            {

                LoadFile(p.Object.File);


            }
            else
            {
                _logger.Info($"Processing '{p.Object.Directory}'");
                _logger.Info("");

                foreach (var file in Directory.GetFiles(p.Object.Directory,"*.pf",SearchOption.AllDirectories))
                {
                    LoadFile(file);
                }
            }
        }

        private static void LoadFile(string pfFile)
        {
            _logger.Warn($"Processing '{pfFile}'");
            _logger.Info("");

            try
            {
                //TODO make these have max length based on length of each count
                string dirFormat = "0000.##";
                string fileFormat = "0000.##";

                var pf = Prefetch.Prefetch.Open(pfFile);

                _logger.Info($"Executable name: {pf.Header.ExecutableFilename}");
                _logger.Info($"Hash: {pf.Header.Hash}");
                _logger.Info($"Version: {pf.Header.Version}");
                _logger.Info("");

                _logger.Info($"Run count: {pf.RunCount}");
                _logger.Warn($"Last run: {pf.LastRunTimes.First()}");
                _logger.Info($"Other run times: {string.Join(",", pf.LastRunTimes.Skip(1))}");

                _logger.Info("");
                _logger.Info("Volume info:");
                _logger.Info("");
                var volnum = 0;
                foreach (var volumeInfo in pf.VolumeInformation)
                {
                    _logger.Info($"#{volnum}: Name: {volumeInfo.DeviceName} Serial: {volumeInfo.SerialNumber} Created: {volumeInfo.CreationTime} Directories: {volumeInfo.DirectoryNames.Count} File references: {volumeInfo.FileReferences.Count}");
                    volnum += 1;
                }
                _logger.Info("");

                _logger.Info($"Directories referenced: {pf.TotalDirectoryCount}");
                _logger.Info("");
                var dirIndex = 0;
                foreach (var volumeInfo in pf.VolumeInformation)
                {
                    foreach (var directoryName in volumeInfo.DirectoryNames)
                    {
                        _logger.Info($"{dirIndex.ToString(dirFormat)}: {directoryName}");
                        dirIndex += 1;
                    }
                }

                _logger.Info("");

                _logger.Info($"Files referenced: {pf.Filenames.Count}");
                _logger.Info("");
                var fileIndex = 0;

                

                foreach (var filename in pf.Filenames)
                {
                    if (filename.Contains(pf.Header.ExecutableFilename))
                    {
                        _logger.Error($"{fileIndex.ToString(fileFormat)}: {filename}");
                    }
                    else
                    {
                        _logger.Info($"{fileIndex.ToString(fileFormat)}: {filename}");
                    }
                    fileIndex += 1;
                }

                _logger.Info("---------------------------------------------------------");

                
            }
            catch (Exception ex)
            {
                _logger.Error($"Error opening '{pfFile}'. Message: {ex.Message}");
            }
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
        public string File { get; set; }
        public string Directory { get; set; }
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
