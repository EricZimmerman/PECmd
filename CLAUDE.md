# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

PECmd (Prefetch Explorer Command Line edition), part of Eric Zimmerman's DFIR tool suite: a CLI (`net472;net10.0`) that parses Windows prefetch files (`.pf`) and shows them on the console or exports them to CSV (with a time line CSV), JSON and HTML. The parsing is in the `Prefetch` NuGet package (its own repository, EricZimmerman/Prefetch); fixes to the format belong there. ERZHelpers does the raw reads of locked files and the Volume Shadow Copies (`--vss`).

Input may be corrupt or hostile: report problems and keep going rather than giving up on the run.

## Commands

All projects use `LangVersion 10`, so avoid newer C# syntax. There is no test project.

```bash
dotnet build PECmd.sln -c Release
dotnet run --project PECmd -f net10.0 -- -d <folder with .pf files> --csv out
```

`Directory.Build.props`: pdbs are embedded in the assemblies (`DebugType` embedded, `EmbedUntrackedSources`) and Release builds set `ContinuousIntegrationBuild` (source paths `/_/...`, not this machine's), so stack traces (Exceptionless) have file and line numbers and no .pdb file ships. Don't add `DebugType` or `DebugSymbols` to project files.

## CLI

All in `PECmd/Program.cs`.
- `Main` returns the exit code, 0 or 1 only. 1 is a failure of the tool: invalid arguments, a missing `-f`/`-d`, a file or folder it can't open, read or list (`_readErrors`: access denied, locked, IO errors, VSC/`--ads` listing, a file `--dedupe` can't hash, a folder under `-d` that can't be read: `ReportUnreadableFolder` logs "Unable to read folder <path>: <reason>" once per folder and the search goes on; net10.0 finds files with `EntryFinder`, net472 with an AlphaFS `ErrorFilter` and without `ContinueOnException`, which stops AlphaFS calling the filter), or output it can't write (`exportErrors`). Damaged files are listed in `_failedFiles` but exit 0.
- `LoadFile` opens a file and lists it in `_failedFiles` when it throws or `ParsingError` is set; `DisplayFile` (without `-q`) only shows it, with the parse problems. The summary's processed count is the files found minus `_failedFiles`, so a file goes in that list once.
- A prefetch file can have no last run times (run count 0); don't assume `LastRunTimes` has any.
- Runs on Linux and macOS too: Windows 10/11 files are decompressed without Windows; `--vss` and `--ads` are Windows only (a warning elsewhere). `#if NET9_0_OR_GREATER` separates the two targets: build both.
- CSV columns: keep the existing ones in order and add new ones at the end.
- `--co` (Issues #259, all CLI tools, opt-in): `ConsoleLog.cs` adds a Serilog file sink, `<yyyyMMddHHmmss>_PECmd_<version>_ConsoleLog.txt` in the first of `--csv`, `--json`, `--html`, with the same `utcNow` as the output names (set before the logger), template `[HH:mm:ss.fff LVL] {Message:l}` (not `:lj`: with `:j` the file sink quotes dates and enums), same formatter and level switch as the console. `DateTimeOffsetFormatter` uses an explicit format when one is given (the templates' `{Timestamp:HH:mm:ss.fff}`) and `--dt` otherwise. Check: with `--debug` the log equals the console minus blank lines.

## Versioning and release

- `<Version>` in `PECmd/PECmd.csproj` (`yyyy.mm.x`), bumped by hand per release; user-facing changes go in `D:\ChangeLog.txt` (PECmd section).
- Release builds are staged, signed and zipped by `Build-ZimmermanTools.ps1` in the get-zimmermantools repository (cloned next to this one). Its program table says what goes in this tool's download zip; a new output file or a change in how the tool is built needs that entry updated.
