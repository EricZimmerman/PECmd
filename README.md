# PECmd

## Command Line Interface

     PECmd version 1.4.0.0
    
     Author: Eric Zimmerman (saericzimmerman@gmail.com)
     https://github.com/EricZimmerman/PECmd
     
             d               Directory to recursively process. Either this or -f is required
             f               File to process. Either this or -d is required
             k               Comma separated list of keywords to highlight in output. By default, 'temp' and 'tmp' are highlighted. Any additional keywords will be added to these.
             o               When specified, save prefetch file bytes to the given path. Useful to look at decompressed Win10 files
             q               Do not dump full details about each file processed. Speeds up processing when using --json or --csv. Default is FALSE
     
             json            Directory to save json representation to.
             jsonf           File name to save JSON formatted results to. When present, overrides default name
             csv             Directory to save CSV results to. Be sure to include the full path in double quotes
             csvf            File name to save CSV formatted results to. When present, overrides default name
             html            Directory to save xhtml formatted results to. Be sure to include the full path in double quotes
             dt              The custom date/time format to use when displaying timestamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss
             mp              When true, display higher precision for timestamps. Default is FALSE
     
             vss             Process all Volume Shadow Copies that exist on drive specified by -f or -d . Default is FALSE
             dedupe          Deduplicate -f or -d & VSCs based on SHA-1. First file found wins. Default is TRUE
             ads             Scan alternate data streams of every file under -d (or the -f file) and parse any prefetch data hidden in them. Default is FALSE

             debug           Show debug information during processing
             trace           Show trace information during processing
     
     Examples: PECmd.exe -f "C:\Temp\CALC.EXE-3FBEF7FD.pf"
               PECmd.exe -f "C:\Temp\CALC.EXE-3FBEF7FD.pf" --json "D:\jsonOutput"
               PECmd.exe -d "C:\Temp" -k "system32, fonts"
               PECmd.exe -d "C:\Temp" --csv "c:\temp" --csvf foo.csv --json c:\temp\json
               PECmd.exe -d "C:\Windows\Prefetch"
               PECmd.exe -d "C:\Windows\Prefetch" --ads --csv "c:\temp"

               Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes

## Alternate Data Stream (ADS) scanning

Use `--ads` to scan the NTFS alternate data streams of files for prefetch data hidden inside them.

When an executable is run from an alternate data stream (for example `notepad.exe` stored as `host.txt:np.exe`), Windows records a prefetch entry whose carrier file has an empty primary stream and the real prefetch tucked into an ADS, e.g. `C:\Windows\Prefetch\HOST.TXT:NP.EXE-F3E0231A.pf`. A normal prefetch scan never looks inside streams, so these are missed. `--ads` finds and parses them.

Behavior:

- `--ads` is opt-in and off by default; without it, PECmd behaves exactly as before.
- With `-d`, every file under the directory is checked (not just `.pf` files), and any stream that parses as a valid prefetch is reported.
- With `-f`, the given file's own alternate data streams are scanned.
- Detection is content based: each stream is parsed as prefetch, so a hidden prefetch is found regardless of the stream's name. Streams that are not prefetch (such as `Zone.Identifier`) are ignored.
- Prefetch recovered from a stream is flagged in the CSV `Note` column as `Prefetch found in ADS`.
- An alternate data stream has no timestamps of its own, so the source Created/Modified/Accessed values shown for an ADS-hosted prefetch are those of the carrier (host) file, which the stream inherits.

Example:

    PECmd.exe -d "C:\Windows\Prefetch" --ads --csv "c:\temp"

## Documentation

If you are running less than Windows 8 you will NOT be able to process Windows 10 prefetch files.

[Windows Prefetch parser in C#](https://binaryforay.blogspot.com/2016/01/windows-prefetch-parser-in-c.html)
[Introducing PECmd!](https://binaryforay.blogspot.com/2016/01/introducing-pecmd.html)
[PECmd v0.6.0.0 released](https://binaryforay.blogspot.com/2016/01/pecmd-v0600-released.html)
[PECmd, LECmd, and JLECmd updated!](https://binaryforay.blogspot.com/2016/03/pecmd-lecmd-and-jlecmd-updated.html)

# Download Eric Zimmerman's Tools

All of Eric Zimmerman's tools can be downloaded [here](https://ericzimmerman.github.io/#!index.md). 

# Special Thanks

Open Source Development funding and support provided by the following contributors: 
- [SANS Institute](http://sans.org/) and [SANS DFIR](http://dfir.sans.org/).
- [Tines](https://www.tines.com/?utm_source=oss&utm_medium=sponsorship&utm_campaign=ericzimmerman)
