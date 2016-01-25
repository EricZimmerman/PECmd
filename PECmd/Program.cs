using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Fclp;
using Fclp.Internals.Extensions;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Targets;
using Prefetch;

namespace PECmd
{
    internal class Program
    {
        private static Logger _logger;

        private static HashSet<string> _keywords;
        private static FluentCommandLineParser<ApplicationArguments> _fluentCommandLineParser;

        private static bool CheckForDotnet46()
        {
            using (
                var ndpKey =
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                        .OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\\"))
            {
                var releaseKey = Convert.ToInt32(ndpKey?.GetValue("Release"));

                return releaseKey >= 393295;
            }
        }

        private static void Main(string[] args)
        {
            SetupNLog();

            _keywords = new HashSet<string> {"temp", "tmp"};

            _logger = LogManager.GetCurrentClassLogger();

            if (!CheckForDotnet46())
            {
                _logger.Warn(".net 4.6 not detected. Please install .net 4.6 and try again.");
                return;
            }

            _fluentCommandLineParser = new FluentCommandLineParser<ApplicationArguments>
            {
                IsCaseSensitive = false
            };

            _fluentCommandLineParser.Setup(arg => arg.File)
                .As('f')
                .WithDescription("File to process. Either this or -d is required");

            _fluentCommandLineParser.Setup(arg => arg.Directory)
                .As('d')
                .WithDescription("Directory to recursively process. Either this or -f is required");

            _fluentCommandLineParser.Setup(arg => arg.Keywords)
                .As('k')
                .WithDescription(
                    "Comma separated list of keywords to highlight in output. By default, 'temp' and 'tmp' are highlighted. Any additional keywords will be added to these.");

            _fluentCommandLineParser.Setup(arg => arg.JsonDirectory)
                .As("json")
                .WithDescription(
                    "Directory to save json representation to. Use --pretty for a more human readable layout");

            _fluentCommandLineParser.Setup(arg => arg.CsvDirectory)
                .As("csv")
                .WithDescription(
                    "File to save CSV (tab separated) results to. Be sure to include the full path in double quotes");

            _fluentCommandLineParser.Setup(arg => arg.JsonPretty)
                .As("pretty")
                .WithDescription(
                    "When exporting to json, use a more human readable layout").SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.LocalTime)
                .As("local")
                .WithDescription(
                    "Display dates using timezone of machine PECmd is running on vs. UTC").SetDefault(false);

            var header =
                $"PECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
                "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
                "\r\nhttps://github.com/EricZimmerman/PECmd";

            var footer = @"Examples: PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf""" + "\r\n\t " +
                         @" PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf"" --json ""D:\jsonOutput"" --jsonpretty" +
                         "\r\n\t " +
                         @" PECmd.exe -d ""C:\Temp"" -k ""system32, fonts""" + "\r\n\t " +
                         @" PECmd.exe -d ""C:\Temp"" --csv ""c:\temp\prefetch_out.csv"" --local" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa" + "\r\n\t " +
//                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa -m 15 -x 22" + "\r\n\t " +
                         @" PECmd.exe -d ""C:\Windows\Prefetch""" + "\r\n\t ";

            _fluentCommandLineParser.SetupHelp("?", "help")
                .WithHeader(header)
                .Callback(text => _logger.Info(text + "\r\n" + footer));

            var result = _fluentCommandLineParser.Parse(args);

            if (result.HelpCalled)
            {
                return;
            }

            if (result.HasErrors)
            {
                _logger.Error("");
                _logger.Error(result.ErrorText);

                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                return;
            }

            if (_fluentCommandLineParser.Object.File.IsNullOrEmpty() &&
                _fluentCommandLineParser.Object.Directory.IsNullOrEmpty())
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                _logger.Warn("Either -f or -d is required. Exiting");
                return;
            }

            if (_fluentCommandLineParser.Object.File.IsNullOrEmpty() == false &&
                !File.Exists(_fluentCommandLineParser.Object.File))
            {
                _logger.Warn($"File '{_fluentCommandLineParser.Object.File}' not found. Exiting");
                return;
            }

            if (_fluentCommandLineParser.Object.Directory.IsNullOrEmpty() == false &&
                !Directory.Exists(_fluentCommandLineParser.Object.Directory))
            {
                _logger.Warn($"Directory '{_fluentCommandLineParser.Object.Directory}' not found. Exiting");
                return;
            }

            if (_fluentCommandLineParser.Object.Keywords?.Length > 0)
            {
                var kws = _fluentCommandLineParser.Object.Keywords.Split(new[] {','},
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (var kw in kws)
                {
                    _keywords.Add(kw.Trim());
                }
            }

            if (_fluentCommandLineParser.Object.CsvDirectory?.Length > 0)
            {
                var dirName = Path.GetDirectoryName(_fluentCommandLineParser.Object.CsvDirectory);

                if (dirName == null)
                {
                    _logger.Warn(
                        $"Couldn't determine directory for '{_fluentCommandLineParser.Object.CsvDirectory}'. Exiting...");
                    return;
                }

                if (Directory.Exists(dirName) == false)
                {
                    _logger.Warn($"Path to '{_fluentCommandLineParser.Object.CsvDirectory}' doesn't exist. Creating...");
                    Directory.CreateDirectory(dirName);
                }
            }


            _logger.Info(header);
            _logger.Info("");
            _logger.Info($"Command line: {string.Join(" ",Environment.GetCommandLineArgs().Skip(1))}");
            _logger.Info("");
            _logger.Info($"Keywords: {string.Join(", ", _keywords)}");
            _logger.Info("");

            if (_fluentCommandLineParser.Object.File?.Length > 0)
            {
                LoadFile(_fluentCommandLineParser.Object.File);
            }
            else
            {
                _logger.Info($"Looking for prefetch files in '{_fluentCommandLineParser.Object.Directory}'");
                _logger.Info("");

                var pfFiles = Directory.GetFiles(_fluentCommandLineParser.Object.Directory, "*.pf",
                    SearchOption.AllDirectories);

                _logger.Info($"Found '{pfFiles.Length}' Prefetch files");

                var sw = new Stopwatch();
                sw.Start();

                foreach (var file in pfFiles)
                {
                    LoadFile(file);
                }

                sw.Stop();
                _logger.Info($"Processing completed in {sw.Elapsed.TotalSeconds:N4} seconds");
            }
        }

        private string GetCsvFormat(IPrefetch pf)
        {
            var sb = new StringBuilder();
            //todo get csv helper or something here to build this intelligently

            //source file name
            //Source created, mod, access
            //TODO one line per for modified, accessed, created of source file?
            //exe name
            //hash
            //size
            //Version
            //Run count
            //last run

            //TODO do i return ONE row for each time it was run, or have a column "Other run times" with the remaining times in them?

            //TODO how to represent volumes? separate fields, or comma separated? same with vol times, etc

            //comma separated list of dirs
            //comma separated list of files

            //do stuff here

            return sb.ToString();
        }

        private static void SaveJson(IPrefetch pf, bool pretty, string outDir)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmss")}_{Path.GetFileName(pf.SourceFilename)}.json";
                var outFile = Path.Combine(outDir, outName);

                _logger.Info("");
                _logger.Info($"Saving json output to '{outFile}'");

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
                .GetCustomAttributes(typeof (DescriptionAttribute), false)
                .SingleOrDefault() as DescriptionAttribute;
            return attribute == null ? value.ToString() : attribute.Description;
        }

        private static IPrefetch LoadFile(string pfFile)
        {
            _logger.Warn($"Processing '{pfFile}'");
            _logger.Info("");

            //TODO REFACTOR TO Prefetch project
            var fi = new FileInfo(pfFile);
            var created = _fluentCommandLineParser.Object.LocalTime
                ? new DateTimeOffset(fi.CreationTime)
                : new DateTimeOffset(fi.CreationTimeUtc);
            var modified = _fluentCommandLineParser.Object.LocalTime
                ? new DateTimeOffset(fi.LastWriteTime)
                : new DateTimeOffset(fi.LastWriteTimeUtc);
            var accessed = _fluentCommandLineParser.Object.LocalTime
                ? new DateTimeOffset(fi.LastAccessTime)
                : new DateTimeOffset(fi.LastAccessTimeUtc);

            _logger.Info($"Created on: {created}");
            _logger.Info($"Modified on: {modified}");
            _logger.Info($"Last accessed on: {accessed}");

            _logger.Info("");

            var sw = new Stopwatch();
            sw.Start();

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
                _logger.Info($"File size (bytes): {pf.Header.FileSize:N0}");
                _logger.Info($"Version: {GetDescriptionFromEnumValue(pf.Header.Version)}");
                _logger.Info("");

         

                _logger.Info($"Run count: {pf.RunCount:N0}");

                var lastRun = pf.LastRunTimes.First();
                if (_fluentCommandLineParser.Object.LocalTime)
                {
                    lastRun = lastRun.ToLocalTime();
                }

                _logger.Warn($"Last run: {lastRun}");

                if (pf.LastRunTimes.Count > 1)
                {
                    var lastRuns = pf.LastRunTimes.Skip(1).ToList();

                    if (_fluentCommandLineParser.Object.LocalTime)
                    {
                        for (var i = 0; i < lastRuns.Count; i++)
                        {
                            lastRuns[0] = lastRuns[0].ToLocalTime();
                        }
                    }
                    var otherRunTimes = string.Join(", ", lastRuns);

                    _logger.Info($"Other run times: {otherRunTimes}");
                }

                _logger.Info("");
                _logger.Info("Volume information:");
                _logger.Info("");
                var volnum = 0;
                foreach (var volumeInfo in pf.VolumeInformation)
                {
                    var localCreate = volumeInfo.CreationTime;
                    if (_fluentCommandLineParser.Object.LocalTime)
                    {
                        localCreate = localCreate.ToLocalTime();
                    }
                    _logger.Info(
                        $"#{volnum}: Name: {volumeInfo.DeviceName} Serial: {volumeInfo.SerialNumber} Created: {localCreate} Directories: {volumeInfo.DirectoryNames.Count:N0} File references: {volumeInfo.FileReferences.Count:N0}");
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

                sw.Stop();

                if (_fluentCommandLineParser.Object.JsonDirectory?.Length > 0)
                {
                    SaveJson(pf, _fluentCommandLineParser.Object.JsonPretty,
                        _fluentCommandLineParser.Object.JsonDirectory);
                }

                _logger.Info("");
                _logger.Info(
                    $"-------------------------------------------- Processed '{pf.SourceFilename}' in {sw.Elapsed.TotalSeconds:N4} seconds --------------------------------------------");
                _logger.Info("\r\n");
                return pf;
            }
            catch (ArgumentNullException)
            {
                _logger.Error(
                    $"Error opening '{pfFile}'.\r\n\r\nThis appears to be a Windows 10 prefetch file. You must be running Windows 8 or higher to decompress Windows 10 prefetch files");
                _logger.Info("");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error opening '{pfFile}'. Message: {ex.Message}");
                _logger.Info("");
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
        public bool LocalTime { get; set; }
        public string CsvDirectory { get; set; }
    }
}