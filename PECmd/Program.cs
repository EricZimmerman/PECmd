#if !NET6_0
using Directory = Alphaleonis.Win32.Filesystem.Directory;
using File = Alphaleonis.Win32.Filesystem.File;
using FileInfo = Alphaleonis.Win32.Filesystem.FileInfo;
using Path = Alphaleonis.Win32.Filesystem.Path;
#else
using Directory = System.IO.Directory;
using File = System.IO.File;
using Path = System.IO.Path;
#endif
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.NamingConventionBinder;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Exceptionless;
using NLog;
using NLog.Config;
using NLog.Targets;
using PECmd.Properties;
using Prefetch;
using RawCopy;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
using Version = Prefetch.Version;

namespace PECmd;

internal class Program
{
    private const string VssDir = @"C:\___vssMount";
    private static Logger _logger;

    private static readonly string _preciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

    private static HashSet<string> _keywords;

    private static List<string> _failedFiles;

    private static List<IPrefetch> _processedFiles;

    private static readonly string Header =
        $"PECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
        "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
        "\r\nhttps://github.com/EricZimmerman/PECmd";

    private static readonly string Footer = @"Examples: PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf""" + "\r\n\t " +
                                            @"   PECmd.exe -f ""C:\Temp\CALC.EXE-3FBEF7FD.pf"" --json ""D:\jsonOutput"" --jsonpretty" +
                                            "\r\n\t " +
                                            @"   PECmd.exe -d ""C:\Temp"" -k ""system32, fonts""" + "\r\n\t " +
                                            @"   PECmd.exe -d ""C:\Temp"" --csv ""c:\temp"" --csvf foo.csv --json c:\temp\json" +
                                            "\r\n\t " +
                                            @"   PECmd.exe -d ""C:\Windows\Prefetch""" + "\r\n\t " +
                                            "\r\n\t" +
                                            "    Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes";

    private static RootCommand _rootCommand;

    private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

    public static bool IsAdministrator()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return true;
        }

        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static async Task Main(string[] args)
    {
        ExceptionlessClient.Default.Startup("x3MPpeQSBUUsXl3DjekRQ9kYjyN3cr5JuwdoOBpZ");

        SetupNLog();

        _keywords = new HashSet<string> { "temp", "tmp" };

        _logger = LogManager.GetCurrentClassLogger();

        _rootCommand = new RootCommand
        {
            new Option<string>(
                "-f",
                "File to process ($MFT | $J | $Boot | $SDS). Required"),

            new Option<string>(
                "-m",
                "$MFT file to use when -f points to a $J file (Use this to resolve parent path in $J CSV output).\r\n"),

            new Option<string>(
                "--json",
                "Directory to save JSON formatted results to. This or --csv required unless --de or --body is specified"),

            new Option<string>(
                "--jsonf",
                "File name to save JSON formatted results to. When present, overrides default name"),

            new Option<string>(
                "--csv",
                "Directory to save CSV formatted results to. This or --json required unless --de or --body is specified"),

            new Option<string>(
                "--csvf",
                "File name to save CSV formatted results to. When present, overrides default name\r\n")
        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        _rootCommand.Handler = CommandHandler.Create(DoWork);

        await _rootCommand.InvokeAsync(args);
    }


    private static void DoWork(string f, string d, string k, string o, bool q, string json, string jsonf, string csv, string csvf, string html, string dt, bool mp, bool vss, bool dedupe, bool debug, bool trace)
    {
        if (f.IsNullOrEmpty() && d.IsNullOrEmpty())
        {
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
            var hc = new HelpContext(helpBld, _rootCommand, Console.Out);

            helpBld.Write(hc);

            _logger.Warn("Either -f or -d is required. Exiting");
            return;
        }

        if (f.IsNullOrEmpty() == false &&
            !File.Exists(f))
        {
            _logger.Warn($"File '{f}' not found. Exiting");
            return;
        }

        if (d.IsNullOrEmpty() == false &&
            !Directory.Exists(d))
        {
            _logger.Warn($"Directory '{d}' not found. Exiting");
            return;
        }

        if (k?.Length > 0)
        {
            var kws = k.ToLowerInvariant().Split(new[] { ',' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (var kw in kws)
            {
                _keywords.Add(kw.Trim());
            }
        }


        _logger.Info(Header);
        _logger.Info("");
        _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}");

        if (IsAdministrator() == false)
        {
            _logger.Fatal("\r\nWarning: Administrator privileges not found!");
        }

        if (debug)
        {
            foreach (var r in LogManager.Configuration.LoggingRules)
            {
                r.EnableLoggingForLevel(LogLevel.Debug);
            }

            LogManager.ReconfigExistingLoggers();
            _logger.Debug("Enabled debug messages...");
        }

        if (trace)
        {
            foreach (var r in LogManager.Configuration.LoggingRules)
            {
                r.EnableLoggingForLevel(LogLevel.Trace);
            }

            LogManager.ReconfigExistingLoggers();
            _logger.Trace("Enabled trace messages...");
        }

        if (vss & (IsAdministrator() == false))
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                vss = false;
                _logger.Warn("--vss is not supported on non-Windows platforms. Disabling...");
            }
        }

        if (vss & (IsAdministrator() == false))
        {
            _logger.Error("--vss is present, but administrator rights not found. Exiting\r\n");
            return;
        }

        _logger.Info("");
        _logger.Info($"Keywords: {string.Join(", ", _keywords)}");
        _logger.Info("");

        if (mp)
        {
            dt = _preciseTimeFormat;
        }

        if (vss)
        {
            string driveLetter;
            if (f.IsEmpty() == false)
            {
                driveLetter = Path.GetPathRoot(Path.GetFullPath(f))
                    .Substring(0, 1);
            }
            else
            {
                driveLetter = Path.GetPathRoot(Path.GetFullPath(d))
                    .Substring(0, 1);
            }

            Helper.MountVss(driveLetter, VssDir);
            Console.WriteLine();
        }

        _processedFiles = new List<IPrefetch>();

        _failedFiles = new List<string>();

        if (f?.Length > 0)
        {
            IPrefetch pf;

            try
            {
                pf = LoadFile(f, q, dt);

                if (pf != null)
                {
                    if (o.IsNullOrEmpty() == false)
                    {
                        try
                        {
                            if (Directory.Exists(Path.GetDirectoryName(o)) ==
                                false)
                            {
                                Directory.CreateDirectory(
                                    Path.GetDirectoryName(o));
                            }

                            PrefetchFile.SavePrefetch(o, pf);
                            _logger.Info($"Saved prefetch bytes to '{o}'");
                        }
                        catch (Exception e)
                        {
                            _logger.Error($"Unable to save prefetch file. Error: {e.Message}");
                        }
                    }

                    _processedFiles.Add(pf);
                }

                if (vss)
                {
                    var vssDirs = Directory.GetDirectories(VssDir);

                    var root = Path.GetPathRoot(Path.GetFullPath(f));
                    var stem = Path.GetFullPath(f).Replace(root, "");

                    foreach (var vssDir in vssDirs)
                    {
                        var newPath = Path.Combine(vssDir, stem);
                        if (File.Exists(newPath))
                        {
                            pf = LoadFile(newPath, q, dt);
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
                    $"Unable to access '{f}'. Are you running as an administrator? Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.Error(
                    $"Error parsing prefetch file '{f}'. Error: {ex.Message}");
            }
        }
        else
        {
            _logger.Info($"Looking for prefetch files in '{d}'");
            _logger.Info("");

            var pfFiles = new List<string>();

#if !NET6_0
        var enumerationFilters = new DirectoryEnumerationFilters();
        enumerationFilters.InclusionFilter = fsei =>
        {
            if (fsei.Extension.ToUpperInvariant() == ".PF")
            {
                if (File.Exists(fsei.FullPath) == false)
                {
                    return false;
                }

                if (fsei.FileSize == 0)
                {
                    _logger.Debug($"Skipping '{fsei.FullPath}' since size is 0");
                    return false;
                }

                if (fsei.FullPath.ToUpperInvariant().Contains("[ROOT]"))
                {
                    _logger.Fatal($"WARNING: FTK Imager detected! Do not use FTK Imager to mount image files as it does not work properly! Use Arsenal Image Mounter instead");
                    return true;
                }

                var fsi = new FileInfo(fsei.FullPath);
                var ads = fsi.EnumerateAlternateDataStreams().Where(t => t.StreamName.Length > 0).ToList();
                if (ads.Count > 0)
                {
                    _logger.Fatal($"WARNING: '{fsei.FullPath}' has at least one Alternate Data Stream:");
                    foreach (var alternateDataStreamInfo in ads)
                    {
                        _logger.Info($"Name: {alternateDataStreamInfo.StreamName}");

                        var s = File.Open(alternateDataStreamInfo.FullPath, FileMode.Open, FileAccess.Read,
                            FileShare.Read, PathFormat.LongFullPath);

                        IPrefetch pf1 = null;

                        try
                        {
                            pf1 = PrefetchFile.Open(s, $"{fsei.FullPath}:{alternateDataStreamInfo.StreamName}");
                        }
                        catch (Exception e)
                        {
                            _logger.Warn($"Could not process '{fsei.FullPath}'. Error: {e.Message}");
                        }

                        _logger.Info(
                            $"---------- Processed '{fsei.FullPath}' ----------");

                        if (pf1 != null)
                        {
                            if (q == false)
                            {
                                DisplayFile(pf1,q,dt);
                            }

                            _processedFiles.Add(pf1);
                        }
                    }

                }

                return true;
            }


            return false;
        };

        enumerationFilters.RecursionFilter = entryInfo => !entryInfo.IsMountPoint && !entryInfo.IsSymbolicLink;

        enumerationFilters.ErrorFilter = (errorCode, errorMessage, pathProcessed) => true;

        var dirEnumOptions =
            DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
            DirectoryEnumerationOptions.SkipReparsePoints | DirectoryEnumerationOptions.ContinueOnException |
            DirectoryEnumerationOptions.BasicSearch;

        var files2 =
            Directory.EnumerateFileSystemEntries(d, dirEnumOptions, enumerationFilters);

#else


            IEnumerable<string> files3;

            var options = new EnumerationOptions
            {
                IgnoreInaccessible = true,
                MatchCasing = MatchCasing.CaseInsensitive,
                RecurseSubdirectories = true,
                AttributesToSkip = 0
            };

            files3 =
                Directory.EnumerateFileSystemEntries(d, "*.pf", options);
#endif


            try
            {
                pfFiles.AddRange(files3);

                if (vss)
                {
                    var vssDirs = Directory.GetDirectories(VssDir);

                    foreach (var vssDir in vssDirs)
                    {
                        var root = Path.GetPathRoot(Path.GetFullPath(d));
                        var stem = Path.GetFullPath(d).Replace(root, "");

                        var target = Path.Combine(vssDir, stem);

                        _logger.Fatal($"Searching 'VSS{target.Replace($"{VssDir}\\", "")}' for prefetch files...");

#if !NET6_0
                    files2 =
                        Directory.EnumerateFileSystemEntries(target, dirEnumOptions, enumerationFilters);

#else
                        IEnumerable<string> files2;

                        var enumerationOptions = new EnumerationOptions
                        {
                            IgnoreInaccessible = true,
                            MatchCasing = MatchCasing.CaseInsensitive,
                            RecurseSubdirectories = true,
                            AttributesToSkip = 0
                        };

                        files2 =
                            Directory.EnumerateFileSystemEntries(target, "*.pf", enumerationOptions);
#endif

                        try
                        {
                            pfFiles.AddRange(files2);
                        }
                        catch (Exception)
                        {
                            _logger.Fatal($"Could not access all files in '{d}'");
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
                    $"Unable to access '{d}'. Are you running as an administrator? Error: {ua.Message}");
                return;
            }
            catch (Exception ex)
            {
                _logger.Error(
                    $"Error getting prefetch files in '{d}'. Error: {ex.Message}");
                return;
            }

            _logger.Info($"\r\nFound {pfFiles.Count:N0} Prefetch files");
            _logger.Info("");

            var sw = new Stopwatch();
            sw.Start();

            var seenHashes = new HashSet<string>();

            foreach (var file in pfFiles)
            {
                if (File.Exists(file) == false)
                {
                    _logger.Warn($"File '{file}' does not seem to exist any more! Skipping");

                    continue;
                }

                if (dedupe)
                {
                    using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        var sha = Helper.GetSha1FromStream(fs, 0);
                        if (seenHashes.Contains(sha))
                        {
                            _logger.Debug($"Skipping '{file}' as a file with SHA-1 '{sha}' has already been processed");
                            continue;
                        }

                        seenHashes.Add(sha);
                    }
                }

                var pf = LoadFile(file, q, dt);

                if (pf != null)
                {
                    _processedFiles.Add(pf);
                }
            }

            sw.Stop();

            if (q)
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
                CsvWriter csvWriter = null;
                StreamWriter streamWriter = null;

                CsvWriter csvTl = null;
                StreamWriter streamWriterTl = null;

                JsConfig.DateHandler = DateHandler.ISO8601;

                StreamWriter streamWriterJson = null;

                var utcNow = DateTimeOffset.UtcNow;

                if (json?.Length > 0)
                {
                    var outName = $"{utcNow:yyyyMMddHHmmss}_PECmd_Output.json";

                    if (Directory.Exists(json) == false)
                    {
                        _logger.Warn(
                            $"'{json} does not exist. Creating...'");
                        Directory.CreateDirectory(json);
                    }

                    if (jsonf.IsNullOrEmpty() == false)
                    {
                        outName = Path.GetFileName(jsonf);
                    }

                    var outFile = Path.Combine(json, outName);

                    _logger.Warn($"Saving json output to '{outFile}'");

                    streamWriterJson = new StreamWriter(outFile);
                    JsConfig.DateHandler = DateHandler.ISO8601;
                }

                if (csv?.Length > 0)
                {
                    var outName = $"{utcNow:yyyyMMddHHmmss}_PECmd_Output.csv";

                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outName = Path.GetFileName(csvf);
                    }

                    var outNameTl = $"{utcNow:yyyyMMddHHmmss}_PECmd_Output_Timeline.csv";
                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outNameTl =
                            $"{Path.GetFileNameWithoutExtension(csvf)}_Timeline{Path.GetExtension(csvf)}";
                    }


                    var outFile = Path.Combine(csv, outName);
                    var outFileTl = Path.Combine(csv, outNameTl);


                    if (Directory.Exists(csv) == false)
                    {
                        _logger.Warn(
                            $"Path to '{csv}' does not exist. Creating...");
                        Directory.CreateDirectory(csv);
                    }

                    _logger.Warn($"CSV output will be saved to '{outFile}'");
                    _logger.Warn($"CSV time line output will be saved to '{outFileTl}'");

                    try
                    {
                        streamWriter = new StreamWriter(outFile);
                        csvWriter = new CsvWriter(streamWriter, CultureInfo.InvariantCulture);

                        csvWriter.WriteHeader(typeof(CsvOut));
                        csvWriter.NextRecord();

                        streamWriterTl = new StreamWriter(outFileTl);
                        csvTl = new CsvWriter(streamWriterTl, CultureInfo.InvariantCulture);

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

                if (html?.Length > 0)
                {
                    if (Directory.Exists(html) == false)
                    {
                        _logger.Warn(
                            $"'{html} does not exist. Creating...'");
                        Directory.CreateDirectory(html);
                    }

                    var outDir = Path.Combine(html,
                        $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_PECmd_Output_for_{html.Replace(@":\", "_").Replace(@"\", "_")}");

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

                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        Resources.directories.Save(Path.Combine(styleDir, "directories.png"));
                        Resources.filesloaded.Save(Path.Combine(styleDir, "filesloaded.png"));
                    }

                    var outFile = Path.Combine(html, outDir,
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

                if (csv.IsNullOrEmpty() == false ||
                    json.IsNullOrEmpty() == false ||
                    html.IsNullOrEmpty() == false)
                {
                    foreach (var processedFile in _processedFiles)
                    {
                        var csvOut = GetCsvFormat(processedFile, dt);

                        var pfname = csvOut.SourceFilename;

                        if (csvOut.SourceFilename.StartsWith(VssDir))
                        {
                            pfname = $"VSS{csvOut.SourceFilename.Replace($"{VssDir}\\", "")}";
                        }

                        csvOut.SourceFilename = pfname;

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
                                t.RunTime = dateTimeOffset.ToString(dt);

                                csvTl?.WriteRecord(t);
                                csvTl?.NextRecord();
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(
                                $"Error getting time line record for '{processedFile.SourceFilename}' to '{csv}'. Error: {ex.Message}");
                        }

                        try
                        {
                            csvWriter?.WriteRecord(csvOut);
                            csvWriter?.NextRecord();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(
                                $"Error writing CSV record for '{processedFile.SourceFilename}' to '{csv}'. Error: {ex.Message}");
                        }

                        //hack
                        if (streamWriterJson != null)
                        {
                            var dtOld = dt;

                            dt = "o";

                            var o1 = GetCsvFormat(processedFile, dt);

                            streamWriterJson.WriteLine(o1.ToJson());

                            dt = dtOld;
                        }


                        //XHTML
                        xml?.WriteStartElement("Container");
                        xml?.WriteElementString("SourceFile", csvOut.SourceFilename);
                        xml?.WriteElementString("SourceCreated", csvOut.SourceCreated);
                        xml?.WriteElementString("SourceModified", csvOut.SourceModified);
                        xml?.WriteElementString("SourceAccessed", csvOut.SourceAccessed);

                        xml?.WriteElementString("LastRun", csvOut.LastRun);

                        xml?.WriteElementString("PreviousRun0", $"{csvOut.PreviousRun0}");
                        xml?.WriteElementString("PreviousRun1", $"{csvOut.PreviousRun1}");
                        xml?.WriteElementString("PreviousRun2", $"{csvOut.PreviousRun2}");
                        xml?.WriteElementString("PreviousRun3", $"{csvOut.PreviousRun3}");
                        xml?.WriteElementString("PreviousRun4", $"{csvOut.PreviousRun4}");
                        xml?.WriteElementString("PreviousRun5", $"{csvOut.PreviousRun5}");
                        xml?.WriteElementString("PreviousRun6", $"{csvOut.PreviousRun6}");

                        xml?.WriteStartElement("ExecutableName");
                        xml?.WriteAttributeString("title",
                            "Note: The name of the executable tracked by the pf file");
                        xml?.WriteString(csvOut.ExecutableName);
                        xml?.WriteEndElement();

                        xml?.WriteElementString("RunCount", $"{csvOut.RunCount}");

                        xml?.WriteStartElement("Size");
                        xml?.WriteAttributeString("title", "Note: The size of the executable in bytes");
                        xml?.WriteString(csvOut.Size);
                        xml?.WriteEndElement();

                        xml?.WriteStartElement("Hash");
                        xml?.WriteAttributeString("title",
                            "Note: The calculated hash for the pf file that should match the hash in the source file name");
                        xml?.WriteString(csvOut.Hash);
                        xml?.WriteEndElement();

                        xml?.WriteStartElement("Version");
                        xml?.WriteAttributeString("title",
                            "Note: The operating system that generated the prefetch file");
                        xml?.WriteString(csvOut.Version);
                        xml?.WriteEndElement();

                        xml?.WriteElementString("Note", csvOut.Note);

                        xml?.WriteElementString("Volume0Name", csvOut.Volume0Name);
                        xml?.WriteElementString("Volume0Serial", csvOut.Volume0Serial);
                        xml?.WriteElementString("Volume0Created", csvOut.Volume0Created);

                        xml?.WriteElementString("Volume1Name", csvOut.Volume1Name);
                        xml?.WriteElementString("Volume1Serial", csvOut.Volume1Serial);
                        xml?.WriteElementString("Volume1Created", csvOut.Volume1Created);


                        xml?.WriteStartElement("Directories");
                        xml?.WriteAttributeString("title",
                            "A comma separated list of all directories accessed by the executable");
                        xml?.WriteString(csvOut.Directories);
                        xml?.WriteEndElement();

                        xml?.WriteStartElement("FilesLoaded");
                        xml?.WriteAttributeString("title",
                            "A comma separated list of all files that were loaded by the executable");
                        xml?.WriteString(csvOut.FilesLoaded);
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

        if (vss)
        {
            if (Directory.Exists(VssDir))
            {
                foreach (var directory in Directory.GetDirectories(VssDir))
                {
                    Directory.Delete(directory);
                }

#if !NET6_0
                Directory.Delete(VssDir, true, true);
#else
                Directory.Delete(VssDir, true);
#endif
            }
        }
    }

// private static object GetPartialDetails(IPrefetch pf, string dt)
//         {
//             var sb = new StringBuilder();
//
//             if (string.IsNullOrEmpty(pf.SourceFilename) == false)
//             {
//                 sb.AppendLine($"Source file name: {pf.SourceFilename}");
//             }
//
//             if (pf.SourceCreatedOn.Year != 1601)
//             {
//                 sb.AppendLine(
//                     $"Accessed on: {pf.SourceCreatedOn.ToUniversalTime().ToString(dt)}");
//             }
//
//             if (pf.SourceModifiedOn.Year != 1601)
//             {
//                 sb.AppendLine(
//                     $"Modified on: {pf.SourceModifiedOn.ToUniversalTime().ToString(dt)}");
//             }
//
//             if (pf.SourceAccessedOn.Year != 1601)
//             {
//                 sb.AppendLine(
//                     $"Last accessed on: {pf.SourceAccessedOn.ToUniversalTime().ToString(dt)}");
//             }
//
//             if (pf.Header != null)
//             {
//                 if (string.IsNullOrEmpty(pf.Header.Signature) == false)
//                 {
//                     sb.AppendLine($"Source file name: {pf.SourceFilename}");
//                 }
//             }
//
//
//             return sb.ToString();
//         }


    private static CsvOut GetCsvFormat(IPrefetch pf, string dt)
    {
        var created = pf.SourceCreatedOn;
        var modified = pf.SourceModifiedOn;
        var accessed = pf.SourceAccessedOn;

        var volDate = string.Empty;
        var volName = string.Empty;
        var volSerial = string.Empty;

        if (pf.VolumeInformation?.Count > 0)
        {
            var vol0Create = pf.VolumeInformation[0].CreationTime;

            volDate = vol0Create.ToString(dt);

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
            var lr = pf.LastRunTimes[0];

            lrTime = lr.ToString(dt);
        }


        var csOut = new CsvOut
        {
            SourceFilename = pf.SourceFilename,
            SourceCreated = created.ToString(dt),
            SourceModified = modified.ToString(dt),
            SourceAccessed = accessed.ToString(dt),
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
            var lrt = pf.LastRunTimes[1];
            csOut.PreviousRun0 = lrt.ToString(dt);
        }

        if (pf.LastRunTimes.Count > 2)
        {
            var lrt = pf.LastRunTimes[2];
            csOut.PreviousRun1 = lrt.ToString(dt);
        }

        if (pf.LastRunTimes.Count > 3)
        {
            var lrt = pf.LastRunTimes[3];
            csOut.PreviousRun2 = lrt.ToString(dt);
        }

        if (pf.LastRunTimes.Count > 4)
        {
            var lrt = pf.LastRunTimes[4];
            csOut.PreviousRun3 = lrt.ToString(dt);
        }

        if (pf.LastRunTimes.Count > 5)
        {
            var lrt = pf.LastRunTimes[5];
            csOut.PreviousRun4 = lrt.ToString(dt);
        }

        if (pf.LastRunTimes.Count > 6)
        {
            var lrt = pf.LastRunTimes[6];
            csOut.PreviousRun5 = lrt.ToString(dt);
        }

        if (pf.LastRunTimes.Count > 7)
        {
            var lrt = pf.LastRunTimes[7];
            csOut.PreviousRun6 = lrt.ToString(dt);
        }

        if (pf.VolumeInformation?.Count > 1)
        {
            var vol1 = pf.VolumeInformation[1].CreationTime;
            csOut.Volume1Created = vol1.ToString(dt);
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
            ?.GetCustomAttributes(typeof(DescriptionAttribute), false)
            .SingleOrDefault() as DescriptionAttribute;
        return attribute?.Description;
    }

    private static void DisplayFile(IPrefetch pf, bool q, string dt)
    {
        if (pf.ParsingError)
        {
            _failedFiles.Add($"'{pf.SourceFilename}' is corrupt and did not parse completely!");
            _logger.Fatal($"'{pf.SourceFilename}' FILE DID NOT PARSE COMPLETELY!\r\n");
        }

        if (q == false)
        {
            if (pf.ParsingError)
            {
                _logger.Fatal("PARTIAL OUTPUT SHOWN BELOW\r\n");
            }


            var created = pf.SourceCreatedOn;
            var modified = pf.SourceModifiedOn;
            var accessed = pf.SourceAccessedOn;

            _logger.Info($"Created on: {created.ToString(dt)}");
            _logger.Info($"Modified on: {modified.ToString(dt)}");
            _logger.Info(
                $"Last accessed on: {accessed.ToString(dt)}");
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


            _logger.Warn($"Last run: {lastRun.ToString(dt)}");

            if (pf.LastRunTimes.Count > 1)
            {
                var lastRuns = pf.LastRunTimes.Skip(1).ToList();


                var otherRunTimes = string.Join(", ",
                    lastRuns.Select(t => t.ToString(dt)));

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


                _logger.Info(
                    $"#{volnum}: Name: {volumeInfo.DeviceName} Serial: {volumeInfo.SerialNumber} Created: {localCreate.ToString(dt)} Directories: {volumeInfo.DirectoryNames.Count:N0} File references: {volumeInfo.FileReferences.Count:N0}");
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


        if (q == false)
        {
            _logger.Info("");
        }


        if (q == false)
        {
            _logger.Info("\r\n");
        }
    }

    private static IPrefetch LoadFile(string pfFile, bool q, string dt)
    {
        var pfname = pfFile;

        if (pfFile.StartsWith(VssDir))
        {
            pfname = $"VSS{pfFile.Replace($"{VssDir}\\", "")}";
        }

        if (q == false)
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

            if (q == false)
            {
                DisplayFile(pf, false, dt);
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

    private static void SetupNLog()
    {
        if (File.Exists(Path.Combine(BaseDirectory, "Nlog.config")))
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
