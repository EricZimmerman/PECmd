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
using Exceptionless.Logging;
using PECmd.Properties;
using Prefetch;
using RawCopy;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
using Version = Prefetch.Version;
#if NET462
using Directory = Alphaleonis.Win32.Filesystem.Directory;
using File = Alphaleonis.Win32.Filesystem.File;
using FileInfo = Alphaleonis.Win32.Filesystem.FileInfo;
using Path = Alphaleonis.Win32.Filesystem.Path;
#else
using Directory = System.IO.Directory;
using File = System.IO.File;
using Path = System.IO.Path;
#endif

namespace PECmd;

internal class Program
{
    private const string VssDir = @"C:\___vssMount";

    private static readonly string PreciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

    private static string ActiveDateTimeFormat;

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

    private static bool IsAdministrator()
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

        _keywords = new HashSet<string> { "temp", "tmp" };

        _rootCommand = new RootCommand
        {
            new Option<string>(
                "-f",
                "File to process. Either this or -d is required"),

            new Option<string>(
                "-d",
                "Directory to recursively process. Either this or -f is required"),

            
            new Option<string>(
                "-k",
                "Comma separated list of keywords to highlight in output. By default, 'temp' and 'tmp' are highlighted. Any additional keywords will be added to these"),

            new Option<string>(
                "-o",
                "When specified, save prefetch file bytes to the given path. Useful to look at decompressed Win10 files"),

            new Option<bool>(
                "-q",
                getDefaultValue:()=>false,
                "Do not dump full details about each file processed. Speeds up processing when using --json or --csv"),
            
            new Option<string>(
                "--json",
                "Directory to save JSON formatted results to. Be sure to include the full path in double quotes"),

            new Option<string>(
                "--jsonf",
                "File name to save JSON formatted results to. When present, overrides default name"),

            new Option<string>(
                "--csv",
                "Directory to save CSV formatted results to. Be sure to include the full path in double quotes"),

            new Option<string>(
                "--csvf",
                "File name to save CSV formatted results to. When present, overrides default name\r\n"),
            
            new Option<string>(
                "--html",
                "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes"),
            
            new Option<string>(
                "--dt",
                getDefaultValue:()=>"yyyy-MM-dd HH:mm:ss",
                "The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options"),

            new Option<bool>(
                "--mp",
                getDefaultValue:()=>false,
                "When true, display higher precision for timestamps"),
            new Option<bool>(
                "--vss",
                getDefaultValue:()=>false,
                "Process all Volume Shadow Copies that exist on drive specified by -f or -d"),
            new Option<bool>(
                "--dedupe",
                getDefaultValue:()=>false,
                "Deduplicate -f or -d & VSCs based on SHA-1. First file found wins"),
            new Option<bool>(
                "--debug",
                getDefaultValue:()=>false,
                "Show debug information during processing"),
            new Option<bool>(
                "--trace",
                getDefaultValue:()=>false,
                "Show trace information during processing"),
        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        _rootCommand.Handler = CommandHandler.Create(DoWork);

        await _rootCommand.InvokeAsync(args);
        
        Log.CloseAndFlush();
    }


    private static void DoWork(string f, string d, string k, string o, bool q, string json, string jsonf, string csv, string csvf, string html, string dt, bool mp, bool vss, bool dedupe, bool debug, bool trace)
    {
        var levelSwitch = new LoggingLevelSwitch();

        ActiveDateTimeFormat = dt;
        
        if (mp)
        {
            ActiveDateTimeFormat = PreciseTimeFormat;
        }
        
        var formatter  =
            new DateTimeOffsetFormatter(CultureInfo.CurrentCulture);
        

        var template = "{Message:lj}{NewLine}{Exception}";

        if (debug)
        {
            levelSwitch.MinimumLevel = LogEventLevel.Debug;
            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }

        if (trace)
        {
            levelSwitch.MinimumLevel = LogEventLevel.Verbose;
            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }
        
        var conf = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: template,formatProvider: formatter)
            .MinimumLevel.ControlledBy(levelSwitch);
      
        Log.Logger = conf.CreateLogger();
        
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Console.WriteLine();
            Log.Fatal("Non-Windows platforms not supported due to the need to load decompression specific Windows libraries! Exiting...");
            Console.WriteLine();
            Environment.Exit(0);
            return;
        }
        
        if (f.IsNullOrEmpty() && d.IsNullOrEmpty())
        {
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
            var hc = new HelpContext(helpBld, _rootCommand, Console.Out);

            helpBld.Write(hc);

            Log.Warning("Either -f or -d is required. Exiting");
            return;
        }

        if (f.IsNullOrEmpty() == false &&
            !File.Exists(f))
        {
            Log.Warning("File {F} not found. Exiting",f);
            return;
        }

        if (d.IsNullOrEmpty() == false &&
            !Directory.Exists(d))
        {
            Log.Warning("Directory {D} not found. Exiting",d);
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

        Log.Information("{Header}",Header);
        Console.WriteLine();
        
        Log.Information("Command line: {Args}",string.Join(" ", Environment.GetCommandLineArgs().Skip(1)));

        if (IsAdministrator() == false)
        {
            Console.WriteLine();
            Log.Information("Warning: Administrator privileges not found!");
        }
      
        if (vss & (IsAdministrator() == false))
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                vss = false;
                Log.Warning("--vss is not supported on non-Windows platforms. Disabling...");
            }
        }

        if (vss & (IsAdministrator() == false))
        {
            Log.Error("--vss is present, but administrator rights not found. Exiting");
            Console.WriteLine();
            return;
        }

        Console.WriteLine();
        Log.Information("Keywords: {Keywords}",string.Join(", ", _keywords));
        Console.WriteLine();


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
            try
            {
                var pf = LoadFile(f, q, ActiveDateTimeFormat);

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
                            Log.Information("Saved prefetch bytes to {O}",o);
                        }
                        catch (Exception e)
                        {
                            Log.Error(e,"Unable to save prefetch file. Error: {Message}",e.Message);
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
                            pf = LoadFile(newPath, q, ActiveDateTimeFormat);
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
                Log.Error(ex,
                    "Unable to access {F}. Are you running as an administrator? Error: {Message}",f,ex.Message);
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "Error parsing prefetch file {F}. Error: {Message}",f,ex.Message);
            }
        }
        else
        {
            Log.Information("Looking for prefetch files in {D}",d);
            Console.WriteLine();

            var pfFiles = new List<string>();

            IEnumerable<string> files2;

#if !NET6_0
        var enumerationFilters = new Alphaleonis.Win32.Filesystem.DirectoryEnumerationFilters();
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
                    Log.Debug("Skipping {FullPath} since size is 0",fsei.FullPath);
                    return false;
                }

                if (fsei.FullPath.ToUpperInvariant().Contains("[ROOT]"))
                {
                    Log.Warning("WARNING: FTK Imager detected! Do not use FTK Imager to mount image files as it does not work properly! Use Arsenal Image Mounter instead");
                    return true;
                }

                var fsi = new FileInfo(fsei.FullPath);
                var ads = fsi.EnumerateAlternateDataStreams().Where(t => t.StreamName.Length > 0).ToList();
                if (ads.Count > 0)
                {
                    Log.Warning("WARNING: {FullPath} has at least one Alternate Data Stream",fsei.FullPath);
                    foreach (var alternateDataStreamInfo in ads)
                    {
                        Log.Information("Name: {StreamName}",alternateDataStreamInfo.StreamName);

                        var s = File.Open(alternateDataStreamInfo.FullPath, FileMode.Open, FileAccess.Read,
                            FileShare.Read, Alphaleonis.Win32.Filesystem.PathFormat.LongFullPath);

                        IPrefetch pf1 = null;

                        try
                        {
                            pf1 = PrefetchFile.Open(s, $"{fsei.FullPath}:{alternateDataStreamInfo.StreamName}");
                        }
                        catch (Exception e)
                        {
                            Log.Warning(e,"Could not process {FullPath}. Error: {Message}",fsei.FullPath,e.Message);
                        }

                        Log.Information("---------- Processed {FullPath} ----------",fsei.FullPath);

                        if (pf1 != null)
                        {
                            if (q == false)
                            {
                                DisplayFile(pf1,q,ActiveDateTimeFormat);
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
            Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.Files | Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.Recursive |
            Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.SkipReparsePoints | Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.ContinueOnException |
            Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.BasicSearch;

        files2 =
            Directory.EnumerateFileSystemEntries(d, dirEnumOptions, enumerationFilters);
        
#else
            
            var options = new EnumerationOptions
            {
                IgnoreInaccessible = true,
                MatchCasing = MatchCasing.CaseInsensitive,
                RecurseSubdirectories = true,
                AttributesToSkip = 0
            };

            files2 = Directory.EnumerateFileSystemEntries(d, "*.pf", options);
#endif


            try
            {
                pfFiles.AddRange(files2);

                if (vss)
                {
                    var vssDirs = Directory.GetDirectories(VssDir);

                    foreach (var vssDir in vssDirs)
                    {
                        var root = Path.GetPathRoot(Path.GetFullPath(d));
                        var stem = Path.GetFullPath(d).Replace(root, "");

                        var target = Path.Combine(vssDir, stem);

                        Log.Information("Searching {Target} for prefetch files...",$"VSS{target.Replace($"{VssDir}\\", "")}");

#if !NET6_0
                    files2 =
                        Directory.EnumerateFileSystemEntries(target, dirEnumOptions, enumerationFilters);

#else

                        var enumerationOptions = new EnumerationOptions
                        {
                            IgnoreInaccessible = true,
                            MatchCasing = MatchCasing.CaseInsensitive,
                            RecurseSubdirectories = true,
                            AttributesToSkip = 0
                        };

                         files2 = Directory.EnumerateFileSystemEntries(target, "*.pf", enumerationOptions);
#endif

                        try
                        {
                            pfFiles.AddRange(files2);
                        }
                        catch (Exception e)
                        {
                            Log.Error(e,"Could not access all files in {D}",d);
                            Console.WriteLine();
                            Log.Error("Rerun the program with Administrator privileges to try again");
                            Console.WriteLine();
                            //Environment.Exit(-1);
                        }
                    }
                }
            }
            catch (UnauthorizedAccessException ua)
            {
                Log.Error(ua,
                    "Unable to access {D}. Are you running as an administrator? Error: {Message}",d,ua.Message);
                return;
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "Error getting prefetch files in {D}. Error: {Message}",d,ex.Message);
                return;
            }

            Console.WriteLine();
            Log.Information("Found {Count:N0} Prefetch files",pfFiles.Count);
            Console.WriteLine();

            var sw = new Stopwatch();
            sw.Start();

            var seenHashes = new HashSet<string>();

            foreach (var file in pfFiles)
            {
                if (File.Exists(file) == false)
                {
                    Log.Warning("File {File} does not seem to exist any more! Skipping",file);

                    continue;
                }

                if (dedupe)
                {
                    using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
                    var sha = Helper.GetSha1FromStream(fs, 0);
                    if (seenHashes.Contains(sha))
                    {
                        Log.Debug("Skipping {File} as a file with SHA-1 {Sha} has already been processed",file,sha);
                        continue;
                    }

                    seenHashes.Add(sha);
                }

                var pf = LoadFile(file, q, ActiveDateTimeFormat);

                if (pf != null)
                {
                    _processedFiles.Add(pf);
                }
            }

            sw.Stop();

            if (q)
            {
                Console.WriteLine();
            }

            Log.Information(
                "Processed {FileCount:N0} out of {TotalFilesCount:N0} files in {TotalSeconds:N4} seconds",pfFiles.Count - _failedFiles.Count,pfFiles.Count,sw.Elapsed.TotalSeconds);

            if (_failedFiles.Count > 0)
            {
                Console.WriteLine();
                Log.Warning("Failed files");
                foreach (var failedFile in _failedFiles)
                {
                    Log.Information("  {FailedFile}",failedFile);
                }
            }
        }

        if (_processedFiles.Count > 0)
        {
            Console.WriteLine();

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
                        Log.Information("{Json} does not exist. Creating...",json);
                        Directory.CreateDirectory(json);
                    }

                    if (jsonf.IsNullOrEmpty() == false)
                    {
                        outName = Path.GetFileName(jsonf);
                    }

                    var outFile = Path.Combine(json, outName);

                    Log.Information("Saving json output to {OutFile}",outFile);

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
                        Log.Information("Path to {Csv} does not exist. Creating...",csv);
                        Directory.CreateDirectory(csv);
                    }

                    Log.Information("CSV output will be saved to {OutFile}",outFile);
                    Log.Information("CSV time line output will be saved to {OutFileTl}",outFileTl);

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
                        Log.Error(ex,
                            "Unable to open {OutFile} for writing. CSV export canceled. Error: {Message}",outFile,ex.Message);
                    }
                }


                XmlTextWriter xml = null;

                if (html?.Length > 0)
                {
                    if (Directory.Exists(html) == false)
                    {
                        Log.Information("{Html} does not exist. Creating...",html);
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

                    Log.Information("Saving HTML output to {OutFile}",outFile);

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
                        var csvOut = GetCsvFormat(processedFile, ActiveDateTimeFormat);

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
                                t.RunTime = dateTimeOffset.ToString(ActiveDateTimeFormat);

                                csvTl?.WriteRecord(t);
                                csvTl?.NextRecord();
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex,
                                "Error getting time line record for {SourceFilename} to {Csv}. Error: {Message}",processedFile.SourceFilename,csv,ex.Message);
                        }

                        try
                        {
                            csvWriter?.WriteRecord(csvOut);
                            csvWriter?.NextRecord();
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex,"Error writing CSV record for {SourceFilename} to {Csv}. Error: {Message}",processedFile.SourceFilename,csv,ex.Message);
                        }

                        //hack
                        if (streamWriterJson != null)
                        {
                            var o1 = GetCsvFormat(processedFile, "o");

                            streamWriterJson.WriteLine(o1.ToJson());
                            
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
                Log.Error(ex,"Error exporting data: {Message}",ex.Message);
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

            volDate = vol0Create.ToString(ActiveDateTimeFormat);

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

            lrTime = lr.ToString(ActiveDateTimeFormat);
        }


        var csOut = new CsvOut
        {
            SourceFilename = pf.SourceFilename,
            SourceCreated = created.ToString(ActiveDateTimeFormat),
            SourceModified = modified.ToString(ActiveDateTimeFormat),
            SourceAccessed = accessed.ToString(ActiveDateTimeFormat),
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
            csOut.PreviousRun0 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.LastRunTimes.Count > 2)
        {
            var lrt = pf.LastRunTimes[2];
            csOut.PreviousRun1 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.LastRunTimes.Count > 3)
        {
            var lrt = pf.LastRunTimes[3];
            csOut.PreviousRun2 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.LastRunTimes.Count > 4)
        {
            var lrt = pf.LastRunTimes[4];
            csOut.PreviousRun3 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.LastRunTimes.Count > 5)
        {
            var lrt = pf.LastRunTimes[5];
            csOut.PreviousRun4 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.LastRunTimes.Count > 6)
        {
            var lrt = pf.LastRunTimes[6];
            csOut.PreviousRun5 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.LastRunTimes.Count > 7)
        {
            var lrt = pf.LastRunTimes[7];
            csOut.PreviousRun6 = lrt.ToString(ActiveDateTimeFormat);
        }

        if (pf.VolumeInformation?.Count > 1)
        {
            var vol1 = pf.VolumeInformation[1].CreationTime;
            csOut.Volume1Created = vol1.ToString(ActiveDateTimeFormat);
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


    private static string GetDescriptionFromEnumValue(Enum value)
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
            _failedFiles.Add($"{pf.SourceFilename} is corrupt and did not parse completely!");
            Log.Warning("{SourceFilename} FILE DID NOT PARSE COMPLETELY!",pf.SourceFilename);
            Console.WriteLine();
        }

        if (q == false)
        {
            if (pf.ParsingError)
            {
                Log.Warning("PARTIAL OUTPUT SHOWN BELOW");
                Console.WriteLine();
            }


            var created = pf.SourceCreatedOn;
            var modified = pf.SourceModifiedOn;
            var accessed = pf.SourceAccessedOn;

            Log.Information("Created on: {CreatedOn}",created);
            Log.Information("Modified on: {Modified}",modified);
            Log.Information("Last accessed on: {Accessed}",accessed);
            Console.WriteLine();

            Log.Information("Executable name: {ExecutableFilename}",pf.Header.ExecutableFilename);
            Log.Information("Hash: {Hash}",pf.Header.Hash);
            Log.Information("File size (bytes): {FileSize:N0}",pf.Header.FileSize);
            Log.Information("Version: {Description}",GetDescriptionFromEnumValue(pf.Header.Version));
            Console.WriteLine();

            Log.Information("Run count: {RunCount:N0}",pf.RunCount);

            var lastRun = pf.LastRunTimes.First();

            Log.Information("Last run: {LastRun}",lastRun);

            if (pf.LastRunTimes.Count > 1)
            {
                var lastRuns = pf.LastRunTimes.Skip(1).ToList();

                var otherRunTimes = string.Join(", ",
                    lastRuns.Select(t => t.ToString(ActiveDateTimeFormat)));

                Log.Information("Other run times: {OtherRunTimes}",otherRunTimes);
                
            }

            Console.WriteLine();
            Log.Information("Volume information:");
            Console.WriteLine();
            var volnum = 0;

            foreach (var volumeInfo in pf.VolumeInformation)
            {
                var localCreate = volumeInfo.CreationTime;

                Log.Information(
                    "#{VolumeNumber}: Name: {DeviceName} Serial: {SerialNumber} Created: {LocalCreate} Directories: {DirectoryNamesCount:N0} File references: {FileReferencesCount:N0}",
                    volnum, volumeInfo.DeviceName, volumeInfo.SerialNumber, localCreate,
                    volumeInfo.DirectoryNames.Count, volumeInfo.FileReferences.Count);
                volnum += 1;
            }

            Console.WriteLine();

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

            Log.Information("Directories referenced: {TotalDirs:N0}",totalDirs);
            Console.WriteLine();
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
                            Log.Fatal("{DirIndex:0#}: {DirectoryName} (Keyword {Kw})",dirIndex,directoryName,true);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Log.Information("{DirIndex:0#}: {DirectoryName}",dirIndex,directoryName);
                    }

                    dirIndex += 1;
                }
            }

            Console.WriteLine();

            Log.Information("Files referenced: {FilenamesCount:N0}",pf.Filenames.Count);
            Console.WriteLine();
            var fileIndex = 0;

            foreach (var filename in pf.Filenames)
            {
                if (filename.EndsWith(pf.Header.ExecutableFilename))
                {
                    Log.Information("{FileIndex:0#}: {Filename} (Executable: {Exe}) ",fileIndex,filename,true);
                }
                else
                {
                    var found = false;
                    foreach (var keyword in _keywords)
                    {
                        if (filename.ToLower().Contains(keyword))
                        {
                            Log.Information("{FileIndex:0#}: {Filename} (Keyword: {Kw})", fileIndex, filename,true);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Log.Information("{FileIndex:0#}: {Filename}",fileIndex,filename);
                    }
                }

                fileIndex += 1;
            }
        }


        if (q == false)
        {
            Console.WriteLine();
        }
        if (q == false)
        {
            Console.WriteLine();
        }
    }

    class DateTimeOffsetFormatter : IFormatProvider, ICustomFormatter
    {
        private readonly IFormatProvider _innerFormatProvider;

        public DateTimeOffsetFormatter(IFormatProvider innerFormatProvider)
        {
            _innerFormatProvider = innerFormatProvider;
        }

        public object GetFormat(Type formatType)
        {
            return formatType == typeof(ICustomFormatter) ? this : _innerFormatProvider.GetFormat(formatType);
        }

        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            if (arg is DateTimeOffset)
            {
                var size = (DateTimeOffset)arg;
                return size.ToString(ActiveDateTimeFormat);
            }

            var formattable = arg as IFormattable;
            if (formattable != null)
            {
                return formattable.ToString(format, _innerFormatProvider);
            }

            return arg.ToString();
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
            Log.Information("Processing {Pfname}",pfname);
            Console.WriteLine();
        }

        var sw = new Stopwatch();
        sw.Start();

        try
        {
            
            var pf = PrefetchFile.Open(pfFile);

            if (pf.ParsingError)
            {
                _failedFiles.Add($"{pfname} is corrupt and did not parse completely!");
                Log.Fatal("{Pfname} FILE DID NOT PARSE COMPLETELY!",pfname);
                Console.WriteLine();
            }

            if (q == false)
            {
                DisplayFile(pf, false, ActiveDateTimeFormat);
            }

            Log.Information("---------- Processed {PfName} in {TotalSeconds:N8} seconds ----------",pfname,sw.Elapsed.TotalSeconds);

            return pf;
        }
        catch (ArgumentNullException an)
        {
            Log.Error("Error opening {Pfname}. This appears to be a Windows 10 prefetch file. You must be running Windows 8 or higher to decompress Windows 10 prefetch files",pfname);
            Console.WriteLine();
            _failedFiles.Add(
                $"{pfname} ==> ({an.Message} (This appears to be a Windows 10 prefetch file. You must be running Windows 8 or higher to decompress Windows 10 prefetch files))");
        }
        catch (Exception ex)
        {
            Log.Error(ex,"Error opening {Pfname}. Message: {Message}",pfname,ex.Message);
            Console.WriteLine();

            _failedFiles.Add($"{pfname} ==> ({ex.Message})");
        }

        return null;
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
