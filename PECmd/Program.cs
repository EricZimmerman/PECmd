using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CsvHelper;
using Fclp;
using Fclp.Internals.Extensions;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Targets;
using Prefetch;
using Version = Prefetch.Version;

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

            _fluentCommandLineParser.Setup(arg => arg.CsvFile)
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

            _fluentCommandLineParser.Setup(arg => arg.Quiet)
                .As('q')
                .WithDescription(
                    "Do not dump full details about each file processed. Speeds up processing when using --json or --csv")
                .SetDefault(false);

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

            if (_fluentCommandLineParser.Object.CsvFile?.Length > 0)
            {
                var dirName = Path.GetDirectoryName(_fluentCommandLineParser.Object.CsvFile);

                if (dirName == null)
                {
                    _logger.Warn(
                        $"Couldn't determine directory for '{_fluentCommandLineParser.Object.CsvFile}'. Exiting...");
                    return;
                }

                if (Directory.Exists(dirName) == false)
                {
                    _logger.Warn($"Path to '{_fluentCommandLineParser.Object.CsvFile}' doesn't exist. Creating...");
                    Directory.CreateDirectory(dirName);
                }
            }


            _logger.Info(header);
            _logger.Info("");
            _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}");
            _logger.Info("");
            _logger.Info($"Keywords: {string.Join(", ", _keywords)}");
            _logger.Info("");

            CsvWriter _csv = null;

            if (_fluentCommandLineParser.Object.CsvFile?.Length > 0)
            {
                _csv = new CsvWriter(new StreamWriter(_fluentCommandLineParser.Object.CsvFile));
                _csv.Configuration.Delimiter = $"{'\t'}";
                _csv.WriteHeader(typeof (CsvOut));
            }

            if (_fluentCommandLineParser.Object.File?.Length > 0)
            {
                IPrefetch pf = null;

                try
                {
                    pf = LoadFile(_fluentCommandLineParser.Object.File);
                }
                catch (UnauthorizedAccessException ex)
                {
                    _logger.Error($"Unable to access '{_fluentCommandLineParser.Object.File}'. Are you running as an administrator?");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error($"Error getting prefetch files in '{_fluentCommandLineParser.Object.Directory}'. Error: {ex.Message}");
                    return;
                }

                if (pf != null && _csv != null)
                {
                    var o = GetCsvFormat(pf);
                    _csv.WriteRecord(o);
                }
            }
            else
            {
                _logger.Info($"Looking for prefetch files in '{_fluentCommandLineParser.Object.Directory}'");
                _logger.Info("");

                string[] pfFiles = null;

                try
                {
                    pfFiles = Directory.GetFiles(_fluentCommandLineParser.Object.Directory, "*.pf",
                    SearchOption.AllDirectories);
                }
                catch (UnauthorizedAccessException)
                {
                    _logger.Error($"Unable to access '{_fluentCommandLineParser.Object.Directory}'. Are you running as an administrator?");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error($"Error getting prefetch files in '{_fluentCommandLineParser.Object.Directory}'. Error: {ex.Message}");
                    return;
                }
                
                _logger.Info($"Found {pfFiles.Length} Prefetch files");
                _logger.Info("");

                var sw = new Stopwatch();
                sw.Start();

                foreach (var file in pfFiles)
                {
                    var pf = LoadFile(file);
                    if (pf != null && _csv != null)
                    {
                        var o = GetCsvFormat(pf);
                        _csv.WriteRecord(o);
                    }
                }

                sw.Stop();
                _logger.Info($"Processed {pfFiles.Length:N0} files in {sw.Elapsed.TotalSeconds:N4} seconds");
            }
        }

        private static CsvOut GetCsvFormat(IPrefetch pf)
        {
            var created = _fluentCommandLineParser.Object.LocalTime
                ? pf.SourceCreatedOn.ToLocalTime()
                : pf.SourceCreatedOn;
            var modified = _fluentCommandLineParser.Object.LocalTime
                ? pf.SourceModifiedOn.ToLocalTime()
                : pf.SourceModifiedOn;
            var accessed = _fluentCommandLineParser.Object.LocalTime
                ? pf.SourceAccessedOn.ToLocalTime()
                : pf.SourceAccessedOn;

            var vol0Create = _fluentCommandLineParser.Object.LocalTime
                ? pf.VolumeInformation[0].CreationTime.ToLocalTime()
                : pf.VolumeInformation[0].CreationTime;

            var lr = _fluentCommandLineParser.Object.LocalTime
            ? pf.LastRunTimes[0].ToLocalTime()
            : pf.LastRunTimes[0];

            var csOut = new CsvOut
            {
                SourceFilename = pf.SourceFilename,
                SourceCreated = created,
                SourceModified = modified,
                SourceAccessed = accessed,
                Hash = pf.Header.Hash,
                ExecutableName = pf.Header.ExecutableFilename,
                Size = pf.Header.FileSize,
                Version = GetDescriptionFromEnumValue(pf.Header.Version),
                RunCount = pf.RunCount,
                Volume0Created = vol0Create,
                Volume0Name = pf.VolumeInformation[0].DeviceName,
                Volume0Serial = pf.VolumeInformation[0].SerialNumber,
                LastRun = lr
            };

            if (pf.LastRunTimes.Count > 1)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[1].ToLocalTime()
                    : pf.LastRunTimes[1];
                csOut.PreviousRun0 = lrt;
            }

            if (pf.LastRunTimes.Count > 2)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[2].ToLocalTime()
                    : pf.LastRunTimes[2];
                csOut.PreviousRun1 = lrt;
            }

            if (pf.LastRunTimes.Count > 3)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[3].ToLocalTime()
                    : pf.LastRunTimes[3];
                csOut.PreviousRun2 = lrt;
            }

            if (pf.LastRunTimes.Count > 4)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[4].ToLocalTime()
                    : pf.LastRunTimes[4];
                csOut.PreviousRun3 = lrt;
            }

            if (pf.LastRunTimes.Count > 5)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[5].ToLocalTime()
                    : pf.LastRunTimes[5];
                csOut.PreviousRun4 = lrt;
            }

            if (pf.LastRunTimes.Count > 6)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[6].ToLocalTime()
                    : pf.LastRunTimes[6];
                csOut.PreviousRun5 = lrt;
            }

            if (pf.LastRunTimes.Count > 7)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[7].ToLocalTime()
                    : pf.LastRunTimes[7];
                csOut.PreviousRun6 = lrt;
            }

            if (pf.VolumeCount > 1)
            {
                var vol1 = _fluentCommandLineParser.Object.LocalTime
                    ? pf.VolumeInformation[1].CreationTime.ToLocalTime()
                    : pf.VolumeInformation[1].CreationTime;
                csOut.Volume1Created = vol1;
                csOut.Volume1Name = pf.VolumeInformation[1].DeviceName;
                csOut.Volume1Serial = pf.VolumeInformation[1].SerialNumber;
            }

            if (pf.VolumeCount > 2)
            {
                csOut.Note = "File contains > 2 volumes! Please inspect output from main program for full details!";
            }

            var sbDirs = new StringBuilder();
            foreach (var volumeInfo in pf.VolumeInformation)
            {
                sbDirs.Append(string.Join(", ", volumeInfo.DirectoryNames));
            }

            csOut.Directories = sbDirs.ToString();

            csOut.FilesLoaded = string.Join(", ", pf.Filenames);

            return csOut;
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

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                var pf = PrefetchFile.Open(pfFile);

                var created = _fluentCommandLineParser.Object.LocalTime
                    ? pf.SourceCreatedOn.ToLocalTime()
                    : pf.SourceCreatedOn;
                var modified = _fluentCommandLineParser.Object.LocalTime
                    ? pf.SourceModifiedOn.ToLocalTime()
                    : pf.SourceModifiedOn;
                var accessed = _fluentCommandLineParser.Object.LocalTime
                    ? pf.SourceAccessedOn.ToLocalTime()
                    : pf.SourceAccessedOn;

                _logger.Info($"Created on: {created}");
                _logger.Info($"Modified on: {modified}");
                _logger.Info($"Last accessed on: {accessed}");
                _logger.Info("");

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
                            lastRuns[i] = lastRuns[i].ToLocalTime();
                        }
                    }
                    var otherRunTimes = string.Join(", ", lastRuns);

                    _logger.Info($"Other run times: {otherRunTimes}");
                }

                if (_fluentCommandLineParser.Object.Quiet == false)
                {
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

                    var totalDirs = pf.TotalDirectoryCount;
                    if (pf.Header.Version == Version.WinXpOrWin2K3)
                    {
                        totalDirs = 0;
                        //this has -1 for total directories, so we have to calculate it
                        foreach (var volumeInfo in pf.VolumeInformation)
                        {
                            totalDirs += volumeInfo.DirectoryNames.Count;

                        }
                    }

                    _logger.Info($"Directories referenced: {totalDirs:N0}");
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
                }


                sw.Stop();

                if (_fluentCommandLineParser.Object.JsonDirectory?.Length > 0)
                {
                    SaveJson(pf, _fluentCommandLineParser.Object.JsonPretty,
                        _fluentCommandLineParser.Object.JsonDirectory);
                }

                _logger.Info("");
                _logger.Info(
                    $"---------- Processed '{pf.SourceFilename}' in {sw.Elapsed.TotalSeconds:N4} seconds ----------");
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

        public sealed class CsvOut
        {
            public string Note { get; set; }
            public string SourceFilename { get; set; }
            public DateTimeOffset SourceCreated { get; set; }
            public DateTimeOffset SourceModified { get; set; }
            public DateTimeOffset SourceAccessed { get; set; }
            public string ExecutableName { get; set; }
            public string Hash { get; set; }
            public int Size { get; set; }
            public string Version { get; set; }
            public int RunCount { get; set; }

            public DateTimeOffset LastRun { get; set; }
            public DateTimeOffset? PreviousRun0 { get; set; }
            public DateTimeOffset? PreviousRun1 { get; set; }
            public DateTimeOffset? PreviousRun2 { get; set; }
            public DateTimeOffset? PreviousRun3 { get; set; }
            public DateTimeOffset? PreviousRun4 { get; set; }
            public DateTimeOffset? PreviousRun5 { get; set; }
            public DateTimeOffset? PreviousRun6 { get; set; }

            public string Volume0Name { get; set; }
            public string Volume0Serial { get; set; }
            public DateTimeOffset Volume0Created { get; set; }

            public string Volume1Name { get; set; }
            public string Volume1Serial { get; set; }
            public DateTimeOffset? Volume1Created { get; set; }

            public string Directories { get; set; }
            public string FilesLoaded { get; set; }
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
        public string CsvFile { get; set; }
        public bool Quiet { get; set; }
    }
}