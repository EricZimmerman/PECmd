using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Xml;
using Alphaleonis.Win32.Filesystem;
using Exceptionless;
using Fclp;
using Fclp.Internals.Extensions;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Targets;
using PECmd.Properties;
using Prefetch;
using RawCopy;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
using Directory = Alphaleonis.Win32.Filesystem.Directory;
using File = Alphaleonis.Win32.Filesystem.File;
using FileInfo = Alphaleonis.Win32.Filesystem.FileInfo;
using Path = Alphaleonis.Win32.Filesystem.Path;
using Version = Prefetch.Version;

namespace PECmd
{
    internal class Program
    {
        private static Logger _logger;

        private static readonly string _preciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

        private static HashSet<string> _keywords;
        private static FluentCommandLineParser<ApplicationArguments> _fluentCommandLineParser;

        private static List<string> _failedFiles;

        private static List<IPrefetch> _processedFiles;

        private const string VssDir = @"C:\___vssMount";

        public static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void Main(string[] args)
        {
            ExceptionlessClient.Default.Startup("x3MPpeQSBUUsXl3DjekRQ9kYjyN3cr5JuwdoOBpZ");

            SetupNLog();

            _keywords = new HashSet<string> {"temp", "tmp"};

            _logger = LogManager.GetCurrentClassLogger();

         

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

            _fluentCommandLineParser.Setup(arg => arg.OutFile)
                .As('o')
                .WithDescription(
                    "When specified, save prefetch file bytes to the given path. Useful to look at decompressed Win10 files");

            _fluentCommandLineParser.Setup(arg => arg.Quiet)
                .As('q')
                .WithDescription(
                    "Do not dump full details about each file processed. Speeds up processing when using --json or --csv. Default is FALSE\r\n")
                .SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.JsonDirectory)
                .As("json")
                .WithDescription(
                    "Directory to save json representation to.");
            _fluentCommandLineParser.Setup(arg => arg.JsonName)
                .As("jsonf")
                .WithDescription("File name to save JSON formatted results to. When present, overrides default name");

            _fluentCommandLineParser.Setup(arg => arg.CsvDirectory)
                .As("csv")
                .WithDescription(
                    "Directory to save CSV results to. Be sure to include the full path in double quotes");
            _fluentCommandLineParser.Setup(arg => arg.CsvName)
                .As("csvf")
                .WithDescription("File name to save CSV formatted results to. When present, overrides default name");


            _fluentCommandLineParser.Setup(arg => arg.xHtmlDirectory)
                .As("html")
                .WithDescription(
                    "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes");
//
//            _fluentCommandLineParser.Setup(arg => arg.JsonPretty)
//                .As("pretty")
//                .WithDescription(
//                    "When exporting to json, use a more human readable layout. Default is FALSE\r\n").SetDefault(false);
//
//       

            _fluentCommandLineParser.Setup(arg => arg.DateTimeFormat)
                .As("dt")
                .WithDescription(
                    "The custom date/time format to use when displaying timestamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss")
                .SetDefault("yyyy-MM-dd HH:mm:ss");

            _fluentCommandLineParser.Setup(arg => arg.PreciseTimestamps)
                .As("mp")
                .WithDescription(
                    "When true, display higher precision for timestamps. Default is FALSE\r\n").SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.Vss)
                .As("vss")
                .WithDescription(
                    "Process all Volume Shadow Copies that exist on drive specified by -f or -d . Default is FALSE")
                .SetDefault(false);
            _fluentCommandLineParser.Setup(arg => arg.Dedupe)
                .As("dedupe")
                .WithDescription(
                    "Deduplicate -f or -d & VSCs based on SHA-1. First file found wins. Default is TRUE\r\n")
                .SetDefault(true);

            _fluentCommandLineParser.Setup(arg => arg.Debug)
                .As("debug")
                .WithDescription("Show debug information during processing").SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.Trace)
                .As("trace")
                .WithDescription("Show trace information during processing").SetDefault(false);


            var header =
                $"PECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
                "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
                "\r\nhttps://github.com/EricZimmerman/PECmd";

            var footer = @"Examples: PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf""" + "\r\n\t " +
                         @" PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf"" --json ""D:\jsonOutput"" --jsonpretty" +
                         "\r\n\t " +
                         @" PECmd.exe -d ""C:\Temp"" -k ""system32, fonts""" + "\r\n\t " +
                         @" PECmd.exe -d ""C:\Temp"" --csv ""c:\temp"" --csvf foo.csv --json c:\temp\json" +
                         "\r\n\t " +
                         @" PECmd.exe -d ""C:\Windows\Prefetch""" + "\r\n\t " +
                         "\r\n\t" +
                         "  Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

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

            if (UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.File) &&
                UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.Directory))
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                _logger.Warn("Either -f or -d is required. Exiting");
                return;
            }

            if (UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.File) == false &&
                !File.Exists(_fluentCommandLineParser.Object.File))
            {
                _logger.Warn($"File '{_fluentCommandLineParser.Object.File}' not found. Exiting");
                return;
            }

            if (UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.Directory) == false &&
                !Directory.Exists(_fluentCommandLineParser.Object.Directory))
            {
                _logger.Warn($"Directory '{_fluentCommandLineParser.Object.Directory}' not found. Exiting");
                return;
            }

            if (_fluentCommandLineParser.Object.Keywords?.Length > 0)
            {
                var kws = _fluentCommandLineParser.Object.Keywords.ToLowerInvariant().Split(new[] {','},
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (var kw in kws)
                {
                    _keywords.Add(kw.Trim());
                }
            }


            _logger.Info(header);
            _logger.Info("");
            _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}");

            if (IsAdministrator() == false)
            {
                _logger.Fatal("\r\nWarning: Administrator privileges not found!");
            }

            if (_fluentCommandLineParser.Object.Debug)
            {
                foreach (var r in LogManager.Configuration.LoggingRules)
                {
                    r.EnableLoggingForLevel(LogLevel.Debug);
                }

                LogManager.ReconfigExistingLoggers();
                _logger.Debug("Enabled debug messages...");
            }

            if (_fluentCommandLineParser.Object.Trace)
            {
                foreach (var r in LogManager.Configuration.LoggingRules)
                {
                    r.EnableLoggingForLevel(LogLevel.Trace);
                }

                LogManager.ReconfigExistingLoggers();
                _logger.Trace("Enabled trace messages...");
            }

            if (_fluentCommandLineParser.Object.Vss & (Helper.IsAdministrator() == false))
            {
                _logger.Error("--vss is present, but administrator rights not found. Exiting\r\n");
                return;
            }

            _logger.Info("");
            _logger.Info($"Keywords: {string.Join(", ", _keywords)}");
            _logger.Info("");

            if (_fluentCommandLineParser.Object.PreciseTimestamps)
            {
                _fluentCommandLineParser.Object.DateTimeFormat = _preciseTimeFormat;
            }

            if (_fluentCommandLineParser.Object.Vss)
            {
                string driveLetter;
                if (_fluentCommandLineParser.Object.File.IsEmpty() == false)
                {
                    driveLetter = Path.GetPathRoot(Path.GetFullPath(_fluentCommandLineParser.Object.File))
                        .Substring(0, 1);
                }
                else
                {
                    driveLetter = Path.GetPathRoot(Path.GetFullPath(_fluentCommandLineParser.Object.Directory))
                        .Substring(0, 1);
                }
                
                Helper.MountVss(driveLetter,VssDir);
                Console.WriteLine();
            }

            _processedFiles = new List<IPrefetch>();

            _failedFiles = new List<string>();

            if (_fluentCommandLineParser.Object.File?.Length > 0)
            {
                IPrefetch pf = null;

                try
                {
                    pf = LoadFile(_fluentCommandLineParser.Object.File);

                    if (pf != null)
                    {
                        if (_fluentCommandLineParser.Object.OutFile.IsNullOrEmpty() == false)
                        {
                            try
                            {
                                if (Directory.Exists(Path.GetDirectoryName(_fluentCommandLineParser.Object.OutFile)) ==
                                    false)
                                {
                                    Directory.CreateDirectory(
                                        Path.GetDirectoryName(_fluentCommandLineParser.Object.OutFile));
                                }

                                PrefetchFile.SavePrefetch(_fluentCommandLineParser.Object.OutFile, pf);
                                _logger.Info($"Saved prefetch bytes to '{_fluentCommandLineParser.Object.OutFile}'");
                            }
                            catch (Exception e)
                            {
                                _logger.Error($"Unable to save prefetch file. Error: {e.Message}");
                            }
                        }
                        
                        _processedFiles.Add(pf);
                    }

                    if (_fluentCommandLineParser.Object.Vss)
                    {
                        var vssDirs = Directory.GetDirectories(VssDir);

                        var root = Path.GetPathRoot(Path.GetFullPath(_fluentCommandLineParser.Object.File));
                        var stem = Path.GetFullPath(_fluentCommandLineParser.Object.File).Replace(root, "");

                        foreach (var vssDir in vssDirs)
                        {
                            var newPath = Path.Combine(vssDir, stem);
                            if (File.Exists(newPath))
                            {
                                pf = LoadFile(newPath);
                                if (pf != null)
                                {
                                    _processedFiles.Add(pf);
                                }
                            }
                        }
                    }



                }
                catch (UnauthorizedAccessException ex)
                {
                    _logger.Error(
                        $"Unable to access '{_fluentCommandLineParser.Object.File}'. Are you running as an administrator? Error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error parsing prefetch file '{_fluentCommandLineParser.Object.File}'. Error: {ex.Message}");
                }
            }
            else
            {
                _logger.Info($"Looking for prefetch files in '{_fluentCommandLineParser.Object.Directory}'");
                _logger.Info("");

                var pfFiles = new List<string>();


                var f = new DirectoryEnumerationFilters();
                f.InclusionFilter = fsei =>
                {
                    if (fsei.Extension.ToUpperInvariant() == ".PF" )
                    {
                        return true;
                    }

                    var fsi = new FileInfo(fsei.FullPath);
                    var ads = fsi.EnumerateAlternateDataStreams().Where(t=>t.StreamName.Length > 0).ToList();
                    if (ads.Count>0)
                    {
                        _logger.Fatal($"WARNING: '{fsei.FullPath}' has at least one Alternate Data Stream:");
                        foreach (var alternateDataStreamInfo in ads)
                        {
                            _logger.Info($"Name: {alternateDataStreamInfo.StreamName}");

                            var s = File.Open(alternateDataStreamInfo.FullPath, FileMode.Open, FileAccess.Read,
                                FileShare.Read, PathFormat.LongFullPath);
                            
                            var pf1 = PrefetchFile.Open(s,$"{fsei.FullPath}:{alternateDataStreamInfo.StreamName}");

                            _logger.Info(
                                $"---------- Processed '{fsei.FullPath}' ----------");
                            
                            if (pf1 != null)
                            {
                                if (_fluentCommandLineParser.Object.Quiet == false)
                                {
                                    DisplayFile(pf1);
                                }
                                _processedFiles.Add(pf1);
                            }
                        }
                        
                    }

                    return false;
                };

                f.RecursionFilter = entryInfo => !entryInfo.IsMountPoint && !entryInfo.IsSymbolicLink;

                f.ErrorFilter = (errorCode, errorMessage, pathProcessed) => true;

                var dirEnumOptions =
                    DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
                    DirectoryEnumerationOptions.SkipReparsePoints | DirectoryEnumerationOptions.ContinueOnException |
                    DirectoryEnumerationOptions.BasicSearch;

                var files2 =
                    Directory.EnumerateFileSystemEntries(_fluentCommandLineParser.Object.Directory, dirEnumOptions, f);


                try
                {
                    pfFiles.AddRange(files2);

                    if (_fluentCommandLineParser.Object.Vss)
                    {
                        var vssDirs = Directory.GetDirectories(VssDir);

                        foreach (var vssDir in vssDirs)
                        {
                            var root = Path.GetPathRoot(Path.GetFullPath(_fluentCommandLineParser.Object.Directory));
                            var stem = Path.GetFullPath(_fluentCommandLineParser.Object.Directory).Replace(root, "");

                            var target = Path.Combine(vssDir, stem);

                            _logger.Fatal($"Searching 'VSS{target.Replace($"{VssDir}\\","")}' for prefetch files...");

                            files2 =
                                Directory.EnumerateFileSystemEntries(target, dirEnumOptions, f);

                            try
                            {
                                pfFiles.AddRange(files2);

                            }
                            catch (Exception)
                            {
                                _logger.Fatal($"Could not access all files in '{_fluentCommandLineParser.Object.Directory}'");
                                _logger.Error("");
                                _logger.Fatal("Rerun the program with Administrator privileges to try again\r\n");
                                //Environment.Exit(-1);
                            }                    
                        }

                    }
                }
                catch (UnauthorizedAccessException ua)
                {
                    _logger.Error(
                        $"Unable to access '{_fluentCommandLineParser.Object.Directory}'. Are you running as an administrator? Error: {ua.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error getting prefetch files in '{_fluentCommandLineParser.Object.Directory}'. Error: {ex.Message}");
                    return;
                }

                _logger.Info($"\r\nFound {pfFiles.Count:N0} Prefetch files");
                _logger.Info("");

                var sw = new Stopwatch();
                sw.Start();

                var seenHashes = new HashSet<string>();

                foreach (var file in pfFiles)
                {

                    if (_fluentCommandLineParser.Object.Dedupe)
                    {
                        using (var fs = new FileStream(file,FileMode.Open,FileAccess.Read))
                        {
                            var sha = Helper.GetSha1FromStream(fs,0);
                            if (seenHashes.Contains(sha))
                            {
                                _logger.Debug($"Skipping '{file}' as a file with SHA-1 '{sha}' has already been processed");
                                continue;
                            }

                            seenHashes.Add(sha);    
                        }
                        
                    }
                    
                    var pf = LoadFile(file);

                    if (pf != null)
                    {
                        _processedFiles.Add(pf);
                    }
                }

                sw.Stop();

                if (_fluentCommandLineParser.Object.Quiet)
                {
                    _logger.Info("");
                }

                _logger.Info(
                    $"Processed {pfFiles.Count - _failedFiles.Count:N0} out of {pfFiles.Count:N0} files in {sw.Elapsed.TotalSeconds:N4} seconds");

                if (_failedFiles.Count > 0)
                {
                    _logger.Info("");
                    _logger.Warn("Failed files");
                    foreach (var failedFile in _failedFiles)
                    {
                        _logger.Info($"  {failedFile}");
                    }
                }
            }

            if (_processedFiles.Count > 0)
            {
                _logger.Info("");

                try
                {
                    CsvWriter csv = null;
                    StreamWriter streamWriter = null;

                    CsvWriter csvTl = null;
                    StreamWriter streamWriterTl = null;

                    JsConfig.DateHandler = DateHandler.ISO8601;

                    StreamWriter streamWriterJson = null;

                    var dt = DateTimeOffset.UtcNow;

                    if (_fluentCommandLineParser.Object.JsonDirectory?.Length > 0)
                    {
                        var outName = $"{dt:yyyyMMddHHmmss}_PECmd_Output.json";
                  
                        if (Directory.Exists(_fluentCommandLineParser.Object.JsonDirectory) == false)
                        {
                            _logger.Warn(
                                $"'{_fluentCommandLineParser.Object.JsonDirectory} does not exist. Creating...'");
                            Directory.CreateDirectory(_fluentCommandLineParser.Object.JsonDirectory);
                        }

                        if (_fluentCommandLineParser.Object.JsonName.IsNullOrEmpty() == false)
                        {
                            outName = Path.GetFileName(_fluentCommandLineParser.Object.JsonName);
                        }

                        var outFile = Path.Combine(_fluentCommandLineParser.Object.JsonDirectory, outName);

                        _logger.Warn($"Saving json output to '{outFile}'");

                        streamWriterJson = new StreamWriter(outFile);
                    }

                    if (_fluentCommandLineParser.Object.CsvDirectory?.Length > 0)
                    {
                        var outName = $"{dt:yyyyMMddHHmmss}_PECmd_Output.csv";
                        
                        if (_fluentCommandLineParser.Object.CsvName.IsNullOrEmpty() == false)
                        {
                            outName = Path.GetFileName(_fluentCommandLineParser.Object.CsvName);
                        }
                        
                        var outNameTl = $"{dt:yyyyMMddHHmmss}_PECmd_Output_Timeline.csv";
                        if (_fluentCommandLineParser.Object.CsvName.IsNullOrEmpty() == false)
                        {
                            outNameTl =
                                $"{Path.GetFileNameWithoutExtension(_fluentCommandLineParser.Object.CsvName)}_Timeline{Path.GetExtension(_fluentCommandLineParser.Object.CsvName)}";
                        }


                        var outFile = Path.Combine(_fluentCommandLineParser.Object.CsvDirectory, outName);
                        var outFileTl = Path.Combine(_fluentCommandLineParser.Object.CsvDirectory, outNameTl);


                        if (Directory.Exists(_fluentCommandLineParser.Object.CsvDirectory) == false)
                        {
                            _logger.Warn(
                                $"Path to '{_fluentCommandLineParser.Object.CsvDirectory}' does not exist. Creating...");
                            Directory.CreateDirectory(_fluentCommandLineParser.Object.CsvDirectory);
                        }

                        _logger.Warn($"CSV output will be saved to '{outFile}'");
                        _logger.Warn($"CSV time line output will be saved to '{outFileTl}'");

                        try
                        {
                            streamWriter = new StreamWriter(outFile);
                            csv = new CsvWriter(streamWriter);

                            csv.WriteHeader(typeof(CsvOut));
                            csv.NextRecord();

                            streamWriterTl = new StreamWriter(outFileTl);
                            csvTl = new CsvWriter(streamWriterTl);

                            csvTl.WriteHeader(typeof(CsvOutTl));
                            csvTl.NextRecord();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(
                                $"Unable to open '{outFile}' for writing. CSV export canceled. Error: {ex.Message}");
                        }
                    }



                    XmlTextWriter xml = null;

                    if (_fluentCommandLineParser.Object.xHtmlDirectory?.Length > 0)
                    {
                        if (Directory.Exists(_fluentCommandLineParser.Object.xHtmlDirectory) == false)
                        {
                            _logger.Warn(
                                $"'{_fluentCommandLineParser.Object.xHtmlDirectory} does not exist. Creating...'");
                            Directory.CreateDirectory(_fluentCommandLineParser.Object.xHtmlDirectory);
                        }

                        var outDir = Path.Combine(_fluentCommandLineParser.Object.xHtmlDirectory,
                            $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_PECmd_Output_for_{_fluentCommandLineParser.Object.xHtmlDirectory.Replace(@":\", "_").Replace(@"\", "_")}");

                        if (Directory.Exists(outDir) == false)
                        {
                            Directory.CreateDirectory(outDir);
                        }

                        var styleDir = Path.Combine(outDir, "styles");
                        if (Directory.Exists(styleDir) == false)
                        {
                            Directory.CreateDirectory(styleDir);
                        }

                        File.WriteAllText(Path.Combine(styleDir, "normalize.css"), Resources.normalize);
                        File.WriteAllText(Path.Combine(styleDir, "style.css"), Resources.style);

                        Resources.directories.Save(Path.Combine(styleDir, "directories.png"));
                        Resources.filesloaded.Save(Path.Combine(styleDir, "filesloaded.png"));

                        var outFile = Path.Combine(_fluentCommandLineParser.Object.xHtmlDirectory, outDir,
                            "index.xhtml");

                        _logger.Warn($"Saving HTML output to '{outFile}'");

                        xml = new XmlTextWriter(outFile, Encoding.UTF8)
                        {
                            Formatting = Formatting.Indented,
                            Indentation = 4
                        };

                        xml.WriteStartDocument();

                        xml.WriteProcessingInstruction("xml-stylesheet", "href=\"styles/normalize.css\"");
                        xml.WriteProcessingInstruction("xml-stylesheet", "href=\"styles/style.css\"");

                        xml.WriteStartElement("document");
                    }

                    if (_fluentCommandLineParser.Object.CsvDirectory.IsNullOrEmpty() == false ||
                        _fluentCommandLineParser.Object.JsonDirectory.IsNullOrEmpty() == false ||
                        _fluentCommandLineParser.Object.xHtmlDirectory.IsNullOrEmpty() == false)
                    {
                        foreach (var processedFile in _processedFiles)
                        {
                            var o = GetCsvFormat(processedFile);

                            var pfname = o.SourceFilename;

                            if (o.SourceFilename.StartsWith(VssDir))
                            {
                                pfname=$"VSS{o.SourceFilename.Replace($"{VssDir}\\", "")}";
                            }

                            o.SourceFilename = pfname;

                            try
                            {
                                foreach (var dateTimeOffset in processedFile.LastRunTimes)
                                {
                                    var t = new CsvOutTl();

                                    var exePath =
                                        processedFile.Filenames.FirstOrDefault(
                                            y => y.EndsWith(processedFile.Header.ExecutableFilename));

                                    if (exePath == null)
                                    {
                                        exePath = processedFile.Header.ExecutableFilename;
                                    }

                                    t.ExecutableName = exePath;
                                    t.RunTime = dateTimeOffset.ToString(_fluentCommandLineParser.Object.DateTimeFormat);

                                    csvTl?.WriteRecord(t);
                                    csvTl?.NextRecord();
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.Error(
                                    $"Error getting time line record for '{processedFile.SourceFilename}' to '{_fluentCommandLineParser.Object.CsvDirectory}'. Error: {ex.Message}");
                            }

                            try
                            {
                                csv?.WriteRecord(o);
                                csv?.NextRecord();
                            }
                            catch (Exception ex)
                            {
                                _logger.Error(
                                    $"Error writing CSV record for '{processedFile.SourceFilename}' to '{_fluentCommandLineParser.Object.CsvDirectory}'. Error: {ex.Message}");
                            }

                          
                            streamWriterJson?.WriteLine(o.ToJson());
                         
                            //XHTML
                            xml?.WriteStartElement("Container");
                            xml?.WriteElementString("SourceFile", o.SourceFilename);
                            xml?.WriteElementString("SourceCreated", o.SourceCreated);
                            xml?.WriteElementString("SourceModified", o.SourceModified);
                            xml?.WriteElementString("SourceAccessed", o.SourceAccessed);

                            xml?.WriteElementString("LastRun", o.LastRun);

                            xml?.WriteElementString("PreviousRun0", $"{o.PreviousRun0}");
                            xml?.WriteElementString("PreviousRun1", $"{o.PreviousRun1}");
                            xml?.WriteElementString("PreviousRun2", $"{o.PreviousRun2}");
                            xml?.WriteElementString("PreviousRun3", $"{o.PreviousRun3}");
                            xml?.WriteElementString("PreviousRun4", $"{o.PreviousRun4}");
                            xml?.WriteElementString("PreviousRun5", $"{o.PreviousRun5}");
                            xml?.WriteElementString("PreviousRun6", $"{o.PreviousRun6}");

                            xml?.WriteStartElement("ExecutableName");
                            xml?.WriteAttributeString("title",
                                "Note: The name of the executable tracked by the pf file");
                            xml?.WriteString(o.ExecutableName);
                            xml?.WriteEndElement();

                            xml?.WriteElementString("RunCount", $"{o.RunCount}");

                            xml?.WriteStartElement("Size");
                            xml?.WriteAttributeString("title", "Note: The size of the executable in bytes");
                            xml?.WriteString(o.Size);
                            xml?.WriteEndElement();

                            xml?.WriteStartElement("Hash");
                            xml?.WriteAttributeString("title",
                                "Note: The calculated hash for the pf file that should match the hash in the source file name");
                            xml?.WriteString(o.Hash);
                            xml?.WriteEndElement();

                            xml?.WriteStartElement("Version");
                            xml?.WriteAttributeString("title",
                                "Note: The operating system that generated the prefetch file");
                            xml?.WriteString(o.Version);
                            xml?.WriteEndElement();

                            xml?.WriteElementString("Note", o.Note);

                            xml?.WriteElementString("Volume0Name", o.Volume0Name);
                            xml?.WriteElementString("Volume0Serial", o.Volume0Serial);
                            xml?.WriteElementString("Volume0Created", o.Volume0Created);

                            xml?.WriteElementString("Volume1Name", o.Volume1Name);
                            xml?.WriteElementString("Volume1Serial", o.Volume1Serial);
                            xml?.WriteElementString("Volume1Created", o.Volume1Created);


                            xml?.WriteStartElement("Directories");
                            xml?.WriteAttributeString("title",
                                "A comma separated list of all directories accessed by the executable");
                            xml?.WriteString(o.Directories);
                            xml?.WriteEndElement();

                            xml?.WriteStartElement("FilesLoaded");
                            xml?.WriteAttributeString("title",
                                "A comma separated list of all files that were loaded by the executable");
                            xml?.WriteString(o.FilesLoaded);
                            xml?.WriteEndElement();

                            xml?.WriteEndElement();
                        }


                        //Close CSV stuff
                        streamWriter?.Flush();
                        streamWriter?.Close();

                        streamWriterTl?.Flush();
                        streamWriterTl?.Close();

                        //Close XML
                        xml?.WriteEndElement();
                        xml?.WriteEndDocument();
                        xml?.Flush();

                        //close json
                        streamWriterJson?.Flush();
                        streamWriterJson?.Close();
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error($"Error exporting data! Error: {ex.Message}");
                }
            }

            if (_fluentCommandLineParser.Object.Vss)
            {
                if (Directory.Exists(VssDir))
                {
                    foreach (var directory in Directory.GetDirectories(VssDir))
                    {
                        Directory.Delete(directory);
                    }
                    Directory.Delete(VssDir,true,true);
                }
            }
        }

        private static object GetPartialDetails(IPrefetch pf)
        {
            var sb = new StringBuilder();

            if (string.IsNullOrEmpty(pf.SourceFilename) == false)
            {
                sb.AppendLine($"Source file name: {pf.SourceFilename}");
            }

            if (pf.SourceCreatedOn.Year != 1601)
            {
                sb.AppendLine(
                    $"Accessed on: {pf.SourceCreatedOn.ToUniversalTime().ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
            }

            if (pf.SourceModifiedOn.Year != 1601)
            {
                sb.AppendLine(
                    $"Modified on: {pf.SourceModifiedOn.ToUniversalTime().ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
            }

            if (pf.SourceAccessedOn.Year != 1601)
            {
                sb.AppendLine(
                    $"Last accessed on: {pf.SourceAccessedOn.ToUniversalTime().ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
            }

            if (pf.Header != null)
            {
                if (string.IsNullOrEmpty(pf.Header.Signature) == false)
                {
                    sb.AppendLine($"Source file name: {pf.SourceFilename}");
                }
            }


            return sb.ToString();
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

            var volDate = string.Empty;
            var volName = string.Empty;
            var volSerial = string.Empty;

            if (pf.VolumeInformation?.Count > 0)
            {
                var vol0Create = _fluentCommandLineParser.Object.LocalTime
                    ? pf.VolumeInformation[0].CreationTime.ToLocalTime()
                    : pf.VolumeInformation[0].CreationTime;

                volDate = vol0Create.ToString(_fluentCommandLineParser.Object.DateTimeFormat);

                if (vol0Create.Year == 1601)
                {
                    volDate = string.Empty;
                }

                volName = pf.VolumeInformation[0].DeviceName;
                volSerial = pf.VolumeInformation[0].SerialNumber;
            }

            var lrTime = string.Empty;

            if (pf.LastRunTimes.Count > 0)
            {
                var lr = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[0].ToLocalTime()
                    : pf.LastRunTimes[0];

                lrTime = lr.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }


            var csOut = new CsvOut
            {
                SourceFilename = pf.SourceFilename,
                SourceCreated = created.ToString(_fluentCommandLineParser.Object.DateTimeFormat),
                SourceModified = modified.ToString(_fluentCommandLineParser.Object.DateTimeFormat),
                SourceAccessed = accessed.ToString(_fluentCommandLineParser.Object.DateTimeFormat),
                Hash = pf.Header.Hash,
                ExecutableName = pf.Header.ExecutableFilename,
                Size = pf.Header.FileSize.ToString(),
                Version = GetDescriptionFromEnumValue(pf.Header.Version),
                RunCount = pf.RunCount.ToString(),
                Volume0Created = volDate,
                Volume0Name = volName,
                Volume0Serial = volSerial,
                LastRun = lrTime,
                ParsingError = pf.ParsingError
            };

            if (pf.LastRunTimes.Count > 1)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[1].ToLocalTime()
                    : pf.LastRunTimes[1];
                csOut.PreviousRun0 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.LastRunTimes.Count > 2)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[2].ToLocalTime()
                    : pf.LastRunTimes[2];
                csOut.PreviousRun1 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.LastRunTimes.Count > 3)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[3].ToLocalTime()
                    : pf.LastRunTimes[3];
                csOut.PreviousRun2 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.LastRunTimes.Count > 4)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[4].ToLocalTime()
                    : pf.LastRunTimes[4];
                csOut.PreviousRun3 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.LastRunTimes.Count > 5)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[5].ToLocalTime()
                    : pf.LastRunTimes[5];
                csOut.PreviousRun4 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.LastRunTimes.Count > 6)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[6].ToLocalTime()
                    : pf.LastRunTimes[6];
                csOut.PreviousRun5 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.LastRunTimes.Count > 7)
            {
                var lrt = _fluentCommandLineParser.Object.LocalTime
                    ? pf.LastRunTimes[7].ToLocalTime()
                    : pf.LastRunTimes[7];
                csOut.PreviousRun6 = lrt.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
            }

            if (pf.VolumeInformation?.Count > 1)
            {
                var vol1 = _fluentCommandLineParser.Object.LocalTime
                    ? pf.VolumeInformation[1].CreationTime.ToLocalTime()
                    : pf.VolumeInformation[1].CreationTime;
                csOut.Volume1Created = vol1.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
                csOut.Volume1Name = pf.VolumeInformation[1].DeviceName;
                csOut.Volume1Serial = pf.VolumeInformation[1].SerialNumber;
            }

            if (pf.VolumeInformation?.Count > 2)
            {
                csOut.Note = "File contains > 2 volumes! Please inspect output from main program for full details!";
            }

            var sbDirs = new StringBuilder();
            if (pf.VolumeInformation != null)
            {
                foreach (var volumeInfo in pf.VolumeInformation)
                {
                    sbDirs.Append(string.Join(", ", volumeInfo.DirectoryNames));
                }
            }
            

            if (pf.ParsingError)
            {
                return csOut;
            }

            csOut.Directories = sbDirs.ToString();

            csOut.FilesLoaded = string.Join(", ", pf.Filenames);

            return csOut;
        }

        

        public static string GetDescriptionFromEnumValue(Enum value)
        {
            var attribute = value.GetType()
                .GetField(value.ToString())
                .GetCustomAttributes(typeof(DescriptionAttribute), false)
                .SingleOrDefault() as DescriptionAttribute;
            return attribute?.Description;
        }

        private static void DisplayFile(IPrefetch pf)
        {
            if (pf.ParsingError)
                {
                    _failedFiles.Add($"'{pf.SourceFilename}' is corrupt and did not parse completely!");
                    _logger.Fatal($"'{pf.SourceFilename}' FILE DID NOT PARSE COMPLETELY!\r\n");
                }

                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    if (pf.ParsingError)
                    {
                        _logger.Fatal("PARTIAL OUTPUT SHOWN BELOW\r\n");
                    }


                    var created = _fluentCommandLineParser.Object.LocalTime
                        ? pf.SourceCreatedOn.ToLocalTime()
                        : pf.SourceCreatedOn;
                    var modified = _fluentCommandLineParser.Object.LocalTime
                        ? pf.SourceModifiedOn.ToLocalTime()
                        : pf.SourceModifiedOn;
                    var accessed = _fluentCommandLineParser.Object.LocalTime
                        ? pf.SourceAccessedOn.ToLocalTime()
                        : pf.SourceAccessedOn;

                    _logger.Info($"Created on: {created.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                    _logger.Info($"Modified on: {modified.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                    _logger.Info(
                        $"Last accessed on: {accessed.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
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

                    _logger.Warn($"Last run: {lastRun.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");

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


                        var otherRunTimes = string.Join(", ",
                            lastRuns.Select(t => t.ToString(_fluentCommandLineParser.Object.DateTimeFormat)));

                        _logger.Info($"Other run times: {otherRunTimes}");
                    }

//
//                if (_fluentCommandLineParser.Object.Quiet == false)
//                {
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
                            $"#{volnum}: Name: {volumeInfo.DeviceName} Serial: {volumeInfo.SerialNumber} Created: {localCreate.ToString(_fluentCommandLineParser.Object.DateTimeFormat)} Directories: {volumeInfo.DirectoryNames.Count:N0} File references: {volumeInfo.FileReferences.Count:N0}");
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


                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    _logger.Info("");
                }

              
                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    _logger.Info("\r\n");
                }

        }

        private static IPrefetch LoadFile(string pfFile)
        {
            var pfname = pfFile;

            if (pfFile.StartsWith(VssDir))
            {
                pfname=$"VSS{pfFile.Replace($"{VssDir}\\", "")}";
            }

            if (_fluentCommandLineParser.Object.Quiet == false)
            {
                _logger.Warn($"Processing '{pfname}'");
                _logger.Info("");
            }

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                var pf = PrefetchFile.Open(pfFile);

                if (pf.ParsingError)
                {
                    _failedFiles.Add($"'{pfname}' is corrupt and did not parse completely!");
                    _logger.Fatal($"'{pfname}' FILE DID NOT PARSE COMPLETELY!\r\n");
                }

                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    DisplayFile(pf);
                }
                
                _logger.Info(
                        $"---------- Processed '{pfname}' in {sw.Elapsed.TotalSeconds:N8} seconds ----------");
                
                return pf;
            }
            catch (ArgumentNullException an)
            {
                _logger.Error(
                    $"Error opening '{pfname}'.\r\n\r\nThis appears to be a Windows 10 prefetch file. You must be running Windows 8 or higher to decompress Windows 10 prefetch files");
                _logger.Info("");
                _failedFiles.Add(
                    $"'{pfname}' ==> ({an.Message} (This appears to be a Windows 10 prefetch file. You must be running Windows 8 or higher to decompress Windows 10 prefetch files))");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error opening '{pfname}'. Message: {ex.Message}");
                _logger.Info("");

                _failedFiles.Add($"'{pfname}' ==> ({ex.Message})");
            }

            return null;
        }

        private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static void SetupNLog()
        {
            if (File.Exists( Path.Combine(BaseDirectory,"Nlog.config")))
            {
                return;
            }

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

        public sealed class CsvOutTl
        {
            public string RunTime { get; set; }
            public string ExecutableName { get; set; }
        }

        public sealed class CsvOut
        {
            public string Note { get; set; }
            public string SourceFilename { get; set; }
            public string SourceCreated { get; set; }
            public string SourceModified { get; set; }
            public string SourceAccessed { get; set; }
            public string ExecutableName { get; set; }
            public string Hash { get; set; }
            public string Size { get; set; }
            public string Version { get; set; }
            public string RunCount { get; set; }

            public string LastRun { get; set; }
            public string PreviousRun0 { get; set; }
            public string PreviousRun1 { get; set; }
            public string PreviousRun2 { get; set; }
            public string PreviousRun3 { get; set; }
            public string PreviousRun4 { get; set; }
            public string PreviousRun5 { get; set; }
            public string PreviousRun6 { get; set; }

            public string Volume0Name { get; set; }
            public string Volume0Serial { get; set; }
            public string Volume0Created { get; set; }

            public string Volume1Name { get; set; }
            public string Volume1Serial { get; set; }
            public string Volume1Created { get; set; }

            public string Directories { get; set; }
            public string FilesLoaded { get; set; }
            public bool ParsingError { get; set; }
        }
    }

    internal class ApplicationArguments
    {
        public string File { get; set; }
        public string Directory { get; set; }
        public string Keywords { get; set; }
        public string JsonDirectory { get; set; }
        public string JsonName { get; set; }
        //public bool JsonPretty { get; set; }
        public bool LocalTime { get; set; }
        public string CsvDirectory { get; set; }
        public string CsvName { get; set; }
        public string OutFile { get; set; }
        public bool Quiet { get; set; }
        public bool Vss { get; set; }
        public bool Dedupe { get; set; }
        public string DateTimeFormat { get; set; }

        public bool PreciseTimestamps { get; set; }

        public string xHtmlDirectory { get; set; }

        public bool Debug { get; set; }
        public bool Trace { get; set; }

    }
}