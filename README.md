# PECmd

Parses Windows prefetch (`.pf`) files, including compressed Windows 10/11 ones: the executable's name, full path and
hash, the run count and last run times, the volumes, and the directories and files loaded. Results are shown on the
console and can be exported to CSV (with a time line of every run), JSON and HTML. It can also process Volume Shadow
Copies and find prefetch data hidden in alternate data streams.

| | |
|---|---|
| **Version** | 2026.10.2 |
| **Download** | [.NET 10](https://download.ericzimmermanstools.com/net/PECmd.zip) (recommended) &middot; [.NET Framework 4.7.2](https://download.ericzimmermanstools.com/PECmd.zip) |
| **All tools** | [ericzimmerman.github.io](https://ericzimmerman.github.io) or [Get-ZimmermanTools](https://download.ericzimmermanstools.com/Get-ZimmermanTools.zip), which downloads them all and keeps them current |
| **Changes** | [ChangeLog.txt](https://download.ericzimmermanstools.com/ChangeLog.txt) |
| **Issues** | [EricZimmerman/PECmd](https://github.com/EricZimmerman/PECmd/issues) or [EricZimmerman/Issues](https://github.com/EricZimmerman/Issues/issues) |

## Requirements

- **.NET 10 build**: the [.NET 10 runtime](https://dotnet.microsoft.com/download/dotnet/10.0). Run `PECmd.exe` on Windows,
  or `dotnet PECmd.dll` on any system with the runtime.
- **.NET Framework build**: .NET Framework 4.7.2 or later (Windows only).

## Command line

```text
Usage:
  PECmd [options]

Options:
  -f <f>           File to process. Either this or -d is required
  -d <d>           Directory to recursively process. Either this or -f is required
  -k <k>           Comma separated list of keywords to highlight in output. By default, 'temp' and 'tmp' are highlighted. Any additional keywords will be added to these
  -o <o>           When specified, save prefetch file bytes to the given path. Useful to look at decompressed Win10 files
  -q               Do not dump full details about each file processed. Speeds up processing when using --json or --csv
  --csv <csv>      Directory to save CSV formatted results to. Be sure to include the full path in double quotes
  --csvf <csvf>    File name to save CSV formatted results to. When present, overrides default name
  --json <json>    Directory to save JSON formatted results to. Be sure to include the full path in double quotes
  --jsonf <jsonf>  File name to save JSON formatted results to. When present, overrides default name
  --html <html>    Directory to save xhtml formatted results to. Be sure to include the full path in double quotes
  --dt <dt>        The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options [default: yyyy-MM-dd HH:mm:ss]
  --mp             When true, display higher precision for timestamps
  --vss            Process all Volume Shadow Copies that exist on drive specified by -f or -d
  --dedupe         Deduplicate -f or -d & VSCs based on SHA-1. First file found wins
  --ads            Scan alternate data streams of every file under -d (or the -f file) and parse any prefetch data found hidden in them
  --debug          Show debug information during processing
  --trace          Show trace information during processing
  -?, -h, --help   Show help and usage information
  --version        Show version information
```

Short options (a single letter) start with one dash, long options with two.

## Examples

```text
PECmd.exe -f "C:\Temp\CALC.EXE-3FBEF7FD.pf"
PECmd.exe -f "C:\Temp\CALC.EXE-3FBEF7FD.pf" --json "D:\jsonOutput"
PECmd.exe -d "C:\Temp" -k "system32, fonts"
PECmd.exe -d "C:\Temp" --csv "c:\temp" --csvf foo.csv --json c:\temp\json
PECmd.exe -d "C:\Windows\Prefetch"
```

## Output

- **Console**: every field of each prefetch file, with the keywords from `-k` (and `temp` and `tmp`) highlighted; `-q`
  leaves out the details. A file that doesn't parse completely is shown with its parse problems.
- **CSV** (`--csv`): two files. `<timestamp>_PECmd_Output.csv` has one row per prefetch file, with the executable's
  `FullPath` and `PathValidated` (whether the path matches the prefetch hash). `<timestamp>_PECmd_Output_Timeline.csv`
  has one row per run time (`RunTime`, `ExecutableName`, `SourceFilename`, `PathValidated`). With `--csvf`, the time line
  file is its name with `_Timeline` added.
- **JSON** (`--json`): `<timestamp>_PECmd_Output.json`, one line per prefetch file, time stamps in ISO 8601 UTC
  (`2024-01-02T03:04:05.6780000Z`) with their full precision.
- **HTML** (`--html`): one folder with an `index.xhtml`.
- **Raw bytes** (`-o`): the prefetch file's bytes, decompressed for Windows 10/11 files.

A prefetch file with no last run times (run count 0) shows "Last run: (none)". A file that doesn't parse completely is
listed once under Failed files.

## Alternate Data Stream (ADS) scanning

Use `--ads` to scan the NTFS alternate data streams of files for prefetch data hidden inside them.

When an executable is run from an alternate data stream (for example `notepad.exe` stored as `host.txt:np.exe`), Windows
records a prefetch entry whose carrier file has an empty primary stream and the real prefetch tucked into an ADS, e.g.
`C:\Windows\Prefetch\HOST.TXT:NP.EXE-F3E0231A.pf`. A normal prefetch scan never looks inside streams, so these are
missed. `--ads` finds and parses them.

- `--ads` is opt-in and off by default; without it, PECmd behaves exactly as before.
- `--ads` needs Windows (NTFS alternate data streams); it is disabled with a warning on other platforms.
- With `-d`, every file under the directory is checked (not just `.pf` files), and any stream that parses as a valid
  prefetch is reported.
- With `-f`, the given file's own alternate data streams are scanned.
- Detection is content based: each stream is parsed as prefetch, so a hidden prefetch is found regardless of the
  stream's name. Streams that are not prefetch (such as `Zone.Identifier`) are ignored.
- Prefetch recovered from a stream is flagged in the CSV `Note` column as `Prefetch found in ADS`.
- An alternate data stream has no timestamps of its own, so the source Created/Modified/Accessed values shown for an
  ADS-hosted prefetch are those of the carrier (host) file, which the stream inherits.

```text
PECmd.exe -d "C:\Windows\Prefetch" --ads --csv "c:\temp"
```

## Linux and macOS

The .NET 10 build runs on Linux and macOS and parses every prefetch version, including compressed Windows 10/11 files.
`--vss` and `--ads` need Windows and are disabled with a warning on other platforms.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success. Damaged prefetch files are reported but are not a failure |
| `1` | The tool failed: invalid arguments, a missing `-f` or `-d` path, output that can't be written, or files or folders it can't read, including VSC folders that can't be fully listed and folders or streams that can't be opened during `--ads`. Folders under `-d` that can't be read are reported (`Unable to read folder <path>: <reason>`) and the rest is still processed |

## What's new in 2026.10

- Runs on Linux and macOS: Windows 10/11 prefetch files are decompressed without Windows.
- New `--ads` option: scan alternate data streams for prefetch data.
- Executable full paths in the exports, validated against the prefetch hash.
- JSON time stamps are ISO 8601 UTC with their full precision (they used the `--dt` format, to the second by default).
- Fixes for file references in version 23 and later files (shifted by one entry), version 30 file metrics (all copies of
  the first entry) and 48 bit MFT entry numbers.
- `--vss` works again, and removes only its own VSC links from `C:\___vssMount` when it finishes.
- Exit code 1 when the tool fails (it was always 0). Error reports include source file and line numbers.

The full list is in the [ChangeLog](https://download.ericzimmermanstools.com/ChangeLog.txt).

## Building

```text
dotnet build PECmd.sln -c Release
```

Builds `net10.0` and `net472`. Prefetch parsing is in the [Prefetch](https://github.com/EricZimmerman/Prefetch) library
([NuGet](https://www.nuget.org/packages/Prefetch)).

## Further reading

- [Windows Prefetch parser in C#](https://binaryforay.blogspot.com/2016/01/windows-prefetch-parser-in-c.html)
- [Introducing PECmd!](https://binaryforay.blogspot.com/2016/01/introducing-pecmd.html)
- [PECmd v0.6.0.0 released](https://binaryforay.blogspot.com/2016/01/pecmd-v0600-released.html)
- [PECmd, LECmd, and JLECmd updated!](https://binaryforay.blogspot.com/2016/03/pecmd-lecmd-and-jlecmd-updated.html)

## Special thanks

Open source development funding and support provided by the [SANS Institute](https://www.sans.org/) and
[SANS DFIR](https://www.sans.org/digital-forensics-incident-response/).

## License

[MIT](LICENSE), Copyright (c) Eric R. Zimmerman 2027.
