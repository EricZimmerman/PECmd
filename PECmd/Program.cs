using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Instrumentation;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Fclp;
using Fclp.Internals.Extensions;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Targets;
using Prefetch;


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

        private static HashSet<string> _keywords;

        static void Main(string[] args)
        {
            SetupNLog();

            _keywords = new HashSet<string> {"temp", "tmp"};

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
              .WithDescription("File to process. Either this or -d is required");

            p.Setup(arg => arg.Directory)
                .As('d')
                .WithDescription("Directory to recursively process. Either this or -f is required");

            p.Setup(arg => arg.Keywords)
    .As('k')
    .WithDescription("Comma separated list of keywords to highlight in output. By default, 'temp' and 'tmp' are highlighted. Any additional keywords will be added to these.");

            p.Setup(arg => arg.JsonDirectory)
                .As("json")
                .WithDescription(
                    "Directory to save json representation to. Use --pretty for a more human readable layout");

            p.Setup(arg => arg.JsonPretty)
                .As("pretty")
                .WithDescription(
                    "When exporting to json, use a more human readable layout").SetDefault(false);

            var header =
             $"PECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
             "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
             "\r\nhttps://github.com/EricZimmerman/PECmd";

            var footer = @"Examples: PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf""" + "\r\n\t " +
                         @" PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf"" --json ""D:\jsonOutput"" --jsonpretty" + "\r\n\t " +
                         @" PECmd.exe -d ""C:\Temp"" -k ""system32, fonts""" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa -m 15 -x 22" + "\r\n\t " +
                         @" PECmd.exe -d ""C:\Windows\Prefetch""" + "\r\n\t ";

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

            if (p.Object.Keywords?.Length > 0)
            {
                var kws = p.Object.Keywords.Split(new[] { ',' },StringSplitOptions.RemoveEmptyEntries);

                foreach (var kw in kws)
                {
                    _keywords.Add(kw.Trim());
                }
            }

            _logger.Info(header);
            _logger.Info("");
            _logger.Info($"Keywords: {string.Join(", ",_keywords)}");
            _logger.Info("");

            if (p.Object.File?.Length > 0)
            {
              var pf=  LoadFile(p.Object.File);
                if (pf != null && p.Object.JsonDirectory?.Length > 0)
                {
                    SaveJson(pf, p.Object.JsonPretty, p.Object.JsonDirectory);
                }
            }
            else
            {
                _logger.Info($"Processing '{p.Object.Directory}'");
                _logger.Info("");

                foreach (var file in Directory.GetFiles(p.Object.Directory,"*.pf",SearchOption.AllDirectories))
                {
                    var pf = LoadFile(file);

                    if (pf != null && p.Object.JsonDirectory?.Length > 0)
                    {
                        SaveJson(pf, p.Object.JsonPretty, p.Object.JsonDirectory);
                    }
                }
            }
        }

        private static void SaveJson(IPrefetch pf, bool pretty, string outDir)
        {
            if (Directory.Exists(outDir) == false)
            {
                Directory.CreateDirectory(outDir);
            }

            var outName = $"{DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmss")}_{Path.GetFileName(pf.SourceFilename)}.json";
            var outFile = Path.Combine(outDir, outName);

            _logger.Info("");
            _logger.Info($"Saving json output to '{outFile}'");

            try
            {
                PrefetchFile.DumpToJson(pf, pretty, outFile);
            }
            catch (Exception ex)
            {
                _logger.Error($"Error exporting json. Error: {ex.Message}");
            }
        }


        public static string GetDescriptionFromEnumValue(Enum value)
        {
            var attribute = value.GetType()
                .GetField(value.ToString())
                .GetCustomAttributes(typeof(DescriptionAttribute), false)
                .SingleOrDefault() as DescriptionAttribute;
            return attribute == null ? value.ToString() : attribute.Description;
        }

        private static IPrefetch LoadFile(string pfFile)
        {
            _logger.Warn($"Processing '{pfFile}'");
            _logger.Info("");

            try
            {
                var pf = PrefetchFile.Open(pfFile);

                var dirString = pf.TotalDirectoryCount.ToString(CultureInfo.InvariantCulture);
                var dd = new string('0', dirString.Length);
                var dirFormat = $"{dd}.##";

                var fString = pf.FileMetricsCount.ToString(CultureInfo.InvariantCulture);
                var ff = new string('0', fString.Length);
                var fileFormat = $"{ff}.##";

                _logger.Info($"Executable name: {pf.Header.ExecutableFilename}");
                _logger.Info($"Hash: {pf.Header.Hash}");
                _logger.Info($"Version: {GetDescriptionFromEnumValue( pf.Header.Version)}");
                _logger.Info("");

                _logger.Info($"Run count: {pf.RunCount:N0}");
                _logger.Warn($"Last run: {pf.LastRunTimes.First()}");

                if (pf.LastRunTimes.Count > 1)
                {
                    _logger.Info($"Other run times: {string.Join(",", pf.LastRunTimes.Skip(1))}");
                }
                
                _logger.Info("");
                _logger.Info("Volume information:");
                _logger.Info("");
                var volnum = 0;
                foreach (var volumeInfo in pf.VolumeInformation)
                {
                    _logger.Info(
                        $"#{volnum}: Name: {volumeInfo.DeviceName} Serial: {volumeInfo.SerialNumber} Created: {volumeInfo.CreationTime} Directories: {volumeInfo.DirectoryNames.Count:N0} File references: {volumeInfo.FileReferences.Count:N0}");
                    volnum += 1;
                }
                _logger.Info("");

                _logger.Info($"Directories referenced: {pf.TotalDirectoryCount:N0}");
                _logger.Info("");
                var dirIndex = 0;
                foreach (var volumeInfo in pf.VolumeInformation)
                {

                    foreach (var directoryName in volumeInfo.DirectoryNames)
                    {
                        var found = false;
                        foreach (var keyword in _keywords)
                        {
                            if (directoryName.ToLower().Contains(keyword))
                            {
                                _logger.Fatal($"{dirIndex.ToString(dirFormat)}: {directoryName}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            _logger.Info($"{dirIndex.ToString(dirFormat)}: {directoryName}");
                        }

                        dirIndex += 1;
                    }
                }

                _logger.Info("");

                _logger.Info($"Files referenced: {pf.Filenames.Count:N0}");
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
                        var found = false;
                        foreach (var keyword in _keywords)
                        {
                            if (filename.ToLower().Contains(keyword))
                            {
                                _logger.Fatal($"{fileIndex.ToString(fileFormat)}: {filename}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            _logger.Info($"{fileIndex.ToString(fileFormat)}: {filename}");
                        }
                    }
                    fileIndex += 1;
                }

                _logger.Info("---------------------------------------------------------");

                return pf;
            }
            catch (ArgumentNullException)
            {
                _logger.Error($"Error opening '{pfFile}'.\r\n\r\nThis appears to be a Windows 10 prefetch file. You must be running Windows 8 or higher to decompress Windows 10 prefetch files");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error opening '{pfFile}'. Message: {ex.Message}");
            }

            return null;
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

        public string Keywords { get; set; }

        public string JsonDirectory { get; set; }
        public bool JsonPretty { get; set; }

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
