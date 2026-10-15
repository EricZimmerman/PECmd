using System;
using System.Collections.Generic;
using System.Text;
using Prefetch;
using PrefetchVersion = Prefetch.Version;

namespace PECmd;

internal static class ExecutablePathResolver
{
    public static ExecutablePathResult Resolve(IPrefetch prefetch)
    {
        if (prefetch?.Header == null ||
            string.IsNullOrWhiteSpace(prefetch.Header.ExecutableFilename) ||
            prefetch.Filenames == null)
        {
            return ExecutablePathResult.NotFound;
        }

        var hasExpectedHash = uint.TryParse(
            prefetch.Header.Hash,
            System.Globalization.NumberStyles.AllowHexSpecifier,
            System.Globalization.CultureInfo.InvariantCulture,
            out var expectedHash);
        var candidatesByPriority = new[]
        {
            new List<string>(),
            new List<string>(),
            new List<string>(),
            new List<string>()
        };

        foreach (var filename in prefetch.Filenames)
        {
            if (!HasExecutableName(filename, prefetch.Header.ExecutableFilename))
            {
                continue;
            }

            var candidates = candidatesByPriority[GetCandidatePriority(filename)];
            if (!candidates.Exists(candidate =>
                    string.Equals(candidate, filename, StringComparison.OrdinalIgnoreCase)))
            {
                candidates.Add(filename);
            }
        }

        var preferredCandidates = Array.Find(candidatesByPriority, candidates => candidates.Count > 0);
        if (preferredCandidates == null)
        {
            return ExecutablePathResult.NotFound;
        }

        if (hasExpectedHash)
        {
            foreach (var candidate in preferredCandidates)
            {
                foreach (var hashPath in GetHashPaths(candidate))
                {
                    var calculatedHash = CalculateHash(hashPath, prefetch.Header.Version);
                    if (calculatedHash == expectedHash)
                    {
                        return new ExecutablePathResult(candidate, true);
                    }
                }
            }
        }

        return new ExecutablePathResult(string.Join(", ", preferredCandidates), false);
    }

    private static bool HasExecutableName(string path, string executableName)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        var lastSeparator = Math.Max(path.LastIndexOf('\\'), path.LastIndexOf('/'));
        var filename = lastSeparator >= 0 ? path.Substring(lastSeparator + 1) : path;

        if (string.Equals(filename, executableName, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return executableName.Length == 29 &&
               filename.StartsWith(executableName, StringComparison.OrdinalIgnoreCase) &&
               GetCandidatePriority(filename) < 3;
    }

    private static int GetCandidatePriority(string path)
    {
        if (path.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (path.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }

        if (path.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        return 3;
    }

    private static IEnumerable<string> GetHashPaths(string path)
    {
        yield return path;

        if (!path.StartsWith(@"\VOLUME{", StringComparison.OrdinalIgnoreCase))
        {
            yield break;
        }

        var volumeNameEnd = path.IndexOf('}');
        if (volumeNameEnd < 0 || volumeNameEnd + 1 >= path.Length)
        {
            yield break;
        }

        var pathWithoutVolume = path.Substring(volumeNameEnd + 1);

        // Windows 10+ stores volume GUID paths in the prefetch file, but hashes the
        // process image's native device path. The device number is not preserved.
        for (var volumeNumber = 1; volumeNumber <= 255; volumeNumber++)
        {
            yield return $@"\DEVICE\HARDDISKVOLUME{volumeNumber}{pathWithoutVolume}";
        }
    }

    private static uint CalculateHash(string path, PrefetchVersion version)
    {
        var pathBytes = Encoding.Unicode.GetBytes(path.ToUpperInvariant());

        return version == PrefetchVersion.WinXpOrWin2K3
            ? CalculateXpHash(pathBytes)
            : CalculateVistaOrNewerHash(pathBytes);
    }

    private static uint CalculateXpHash(IEnumerable<byte> pathBytes)
    {
        uint hash = 0;

        unchecked
        {
            foreach (var value in pathBytes)
            {
                hash = hash * 37 + value;
            }

            hash *= 314159269;

            if (hash > 0x80000000)
            {
                hash = 0 - hash;
            }

            hash %= 1000000007;
        }

        return hash;
    }

    private static uint CalculateVistaOrNewerHash(IEnumerable<byte> pathBytes)
    {
        uint hash = 314159;

        unchecked
        {
            foreach (var value in pathBytes)
            {
                hash = hash * 37 + value;
            }
        }

        return hash;
    }
}

internal sealed class ExecutablePathResult
{
    public static readonly ExecutablePathResult NotFound = new(string.Empty, false);

    public ExecutablePathResult(string fullPath, bool pathValidated)
    {
        FullPath = fullPath;
        PathValidated = pathValidated;
    }

    public string FullPath { get; }
    public bool PathValidated { get; }
}
