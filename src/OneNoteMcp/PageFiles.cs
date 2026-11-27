using System.Text;
using System.Text.RegularExpressions;
using ModelContextProtocol;

namespace OneNoteMcp;

/// <summary>
/// Checks, reads, and writes the files agents name in <c>inputPath</c> and <c>outputPath</c>.
/// </summary>
/// <remarks>
/// - Checks run before any OneNote call, so a bad path fails fast and leaves nothing behind.
/// - File errors become <see cref="McpException"/>, so the agent sees the reason.
/// </remarks>
public static partial class PageFiles
{
    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    /// <summary>Rejects paths the tools can't safely read or write.</summary>
    /// <param name="parameter">The tool parameter that holds the path, for error messages</param>
    /// <remarks>
    /// - A relative path would resolve against the server's folder, not the agent's.
    /// - NUL and other device names become <c>\\.\</c> paths once resolved.
    /// </remarks>
    public static void Check(string path, string parameter)
    {
        if (!Path.IsPathFullyQualified(path))
        {
            throw new McpException($"{parameter} must be absolute, such as C:\\work\\page.md, not \"{path}\"");
        }

        string full;
        try
        {
            full = Path.GetFullPath(path);
        }
        catch (Exception e) when (e is ArgumentException or NotSupportedException)
        {
            throw new McpException($"{parameter} isn't a valid path: {e.Message}", e);
        }

        if (IsDevicePath(path) || IsDevicePath(full))
        {
            throw new McpException($"{parameter} must be a file path, not a device path: \"{path}\"");
        }

        var name = Path.GetFileName(path);
        if (name.Length == 0 || Path.GetDirectoryName(path) is null || Directory.Exists(path))
        {
            throw new McpException($"{parameter} must name a file, not a folder: \"{path}\"");
        }

        CheckName(name, parameter);
    }

    /// <summary>Reads a UTF-8 file, or UTF-16 with a byte-order mark.</summary>
    public static string Read(string path)
    {
        string text;
        try
        {
            text = File.ReadAllText(path, StrictUtf8);
        }
        catch (DecoderFallbackException e)
        {
            throw new McpException($"\"{path}\" isn't UTF-8. Save it as UTF-8 and try again.", e);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            throw new McpException($"Can't read \"{path}\": {e.Message}", e);
        }

        // UTF-16 without a byte-order mark reads as valid UTF-8 with a NUL between characters
        if (text.Contains('\0'))
        {
            throw new McpException($"\"{path}\" has a null character. It may be UTF-16 without a byte-order mark: save it as UTF-8.");
        }

        return string.IsNullOrWhiteSpace(text) ? throw new McpException($"\"{path}\" is empty") : text;
    }

    /// <summary>Reads a local image that a page's Markdown or HTML names.</summary>
    /// <param name="baseFolder">The folder relative paths resolve against: the input file's. Null for inline content, which needs absolute paths.</param>
    /// <returns>Null for a web or data URL, a missing or unsafe path, or a file that can't be read, so the image stays a link</returns>
    public static byte[]? ReadImage(string src, string? baseFolder)
    {
        string path;
        if (Uri.TryCreate(src, UriKind.Absolute, out var uri))
        {
            if (!uri.IsFile)
            {
                return null;
            }

            path = uri.LocalPath;
        }
        else if (baseFolder is not null)
        {
            // Saved pages percent-encode their image links
            path = Path.Combine(baseFolder, Uri.UnescapeDataString(src).Replace('/', '\\'));
        }
        else
        {
            return null;
        }

        try
        {
            Check(path, "image path");
            return File.ReadAllBytes(path);
        }
        catch (Exception e) when (e is McpException or IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>Writes UTF-8 without a byte-order mark, creating missing folders.</summary>
    /// <returns>The full path written</returns>
    public static string Write(string path, string content) => Write(path, full => File.WriteAllText(full, content));

    /// <summary>Writes bytes, creating missing folders.</summary>
    /// <returns>The full path written</returns>
    public static string Write(string path, byte[] content) => Write(path, full => File.WriteAllBytes(full, content));

    private static string Write(string path, Action<string> write)
    {
        try
        {
            var full = Path.GetFullPath(path);
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            write(full);
            return full;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            throw new McpException($"Can't save \"{path}\": {e.Message}", e);
        }
    }

    private static bool IsDevicePath(string path) => path.StartsWith(@"\\.\") || path.StartsWith(@"\\?\");

    /// <summary>Rejects file names Windows changes, can't hold, or reserves for devices.</summary>
    /// <remarks>
    /// - Windows drops a trailing dot or space, so the file would land under another name.
    /// - ':' would write to an NTFS alternate data stream.
    /// - CON, AUX, COM1 and the like open devices, or make files many tools can't open or delete.
    /// </remarks>
    private static void CheckName(string name, string parameter)
    {
        if (name.EndsWith('.') || name.EndsWith(' '))
        {
            throw new McpException($"{parameter}'s file name can't end with a dot or space: \"{name}\"");
        }

        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new McpException($"{parameter} has characters a file name can't hold: \"{name}\"");
        }

        if (ReservedName().IsMatch(Path.GetFileNameWithoutExtension(name)))
        {
            throw new McpException($"{parameter} uses a name Windows reserves for devices: \"{name}\"");
        }
    }

    [GeneratedRegex(@"^(CON|CONIN\$|CONOUT\$|PRN|AUX|NUL|COM[0-9\u00B9\u00B2\u00B3]|LPT[0-9\u00B9\u00B2\u00B3])$", RegexOptions.IgnoreCase)]
    private static partial Regex ReservedName();
}
