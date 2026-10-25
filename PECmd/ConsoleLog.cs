using System;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace PECmd;

/// <summary>--co (Issues #259): the console messages are also saved to &lt;yyyyMMddHHmmss&gt;_PECmd_&lt;version&gt;_ConsoleLog.txt in the output directory</summary>
internal static class ConsoleLog
{
    /// <summary>The console's template with time and level. {Message:l}, not :lj: the file sink writes values other than strings as JSON
    /// with :j (dates as quoted ISO 8601 instead of the --dt format)</summary>
    public const string Template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:l}{NewLine}{Exception}";

    /// <summary>The log file in the first output directory given (the time is UTC)</summary>
    /// <returns>null when no output directory was given or the path is not valid (the tool reports that itself)</returns>
    public static string GetPath(DateTimeOffset utcNow, params string[] outputDirectories)
    {
        foreach (var dir in outputDirectories)
        {
            if (string.IsNullOrWhiteSpace(dir))
            {
                continue;
            }

            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version!.ToString(3);
                return Path.Combine(Path.GetFullPath(dir), $"{utcNow:yyyyMMddHHmmss}_PECmd_{version}_ConsoleLog.txt");
            }
            catch (Exception)
            {
                return null;
            }
        }

        return null;
    }

    /// <summary>The tool's formatter for the file, except for a time stamp with a format of its own: the tool's formatter gives every
    /// DateTimeOffset the --dt format, which would replace the template's {Timestamp:HH:mm:ss.fff}</summary>
    public static IFormatProvider Formatter(IFormatProvider toolFormatter)
    {
        return new TimestampFormatter(toolFormatter);
    }

    private sealed class TimestampFormatter : IFormatProvider, ICustomFormatter
    {
        private readonly IFormatProvider _toolFormatter;

        public TimestampFormatter(IFormatProvider toolFormatter)
        {
            _toolFormatter = toolFormatter;
        }

        public object GetFormat(Type formatType)
        {
            return formatType == typeof(ICustomFormatter) ? this : _toolFormatter.GetFormat(formatType);
        }

        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            if (arg is DateTimeOffset dto && string.IsNullOrEmpty(format) == false)
            {
                return dto.ToString(format, CultureInfo.InvariantCulture);
            }

            if (_toolFormatter.GetFormat(typeof(ICustomFormatter)) is ICustomFormatter custom)
            {
                return custom.Format(format, arg, _toolFormatter);
            }

            return arg is IFormattable formattable ? formattable.ToString(format, _toolFormatter) : arg?.ToString();
        }
    }
}
