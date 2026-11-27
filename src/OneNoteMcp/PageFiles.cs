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

    /// <summary>Files larger than this stay links, so a page can't make the server read a huge file.</summary>
    public const long MaxImageBytes = 50L * 1024 * 1024;

    /// <summary>Reads a local image that a page's Markdown or HTML names.</summary>
    /// <param name="baseFolder">The input file's folder, or null for content the agent passed inline.</param>
    /// <returns>Null for anything that should stay a link: a web or data URL, a disallowed or missing path, or a file that's too large or can't be read</returns>
    /// <remarks>
    /// - Files can come from the web, so their images must sit inside the file's folder. Only inline content may use absolute paths.
    /// - Network locations are never read, UNC or mapped drive: opening one sends the user's Windows credentials to that host.
    /// - Neither are paths through a symbolic link or junction, which could point anywhere.
    /// - These checks run on the final full path, before any call that touches the file.
    /// </remarks>
    public static byte[]? ReadImage(string src, string? baseFolder)
    {
        try
        {
            if (ImagePath(src, baseFolder) is not { } candidate || Path.GetFullPath(candidate) is not { } path
                || !IsLocal(path) || GoesThroughLink(path))
            {
                return null;
            }

            Check(path, "image path");
            var file = new FileInfo(path);
            return file.Exists && file.Length <= MaxImageBytes ? File.ReadAllBytes(path) : null;
        }
        catch (Exception e) when (e is McpException or IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or UriFormatException)
        {
            return null;
        }
    }

    /// <returns>The full path of a local image the content may embed, or null</returns>
    private static string? ImagePath(string src, string? baseFolder)
    {
        // Saved pages percent-encode their image links
        var decoded = Uri.UnescapeDataString(src);
        if (IsNetworkPath(src) || IsNetworkPath(decoded))
        {
            return null;
        }

        var isAbsolute = Uri.TryCreate(src, UriKind.Absolute, out var uri);
        if (baseFolder is null)
        {
            if (isAbsolute)
            {
                return uri!.IsFile && !uri.IsUnc ? uri.LocalPath : null;
            }

            return Path.IsPathFullyQualified(decoded) ? decoded : null;
        }

        if (isAbsolute || Path.IsPathRooted(decoded))
        {
            return null;
        }

        // A drive root like C:\ already ends in a separator
        var folder = Path.GetFullPath(baseFolder);
        folder = Path.EndsInDirectorySeparator(folder) ? folder : folder + Path.DirectorySeparatorChar;
        var full = Path.GetFullPath(Path.Combine(folder, decoded.Replace('/', '\\')));
        return full.StartsWith(folder, StringComparison.OrdinalIgnoreCase) ? full : null;
    }

    /// <summary>A UNC path in any spelling, including protocol-relative URLs like <c>//host/a.png</c>.</summary>
    private static bool IsNetworkPath(string path) =>
        path.Length >= 2 && path[0] is '\\' or '/' && path[1] is '\\' or '/';

    /// <summary>A full path on a local drive: not UNC, not a device, not a mapped network drive.</summary>
    private static bool IsLocal(string fullPath) =>
        !IsNetworkPath(fullPath) && !IsDevicePath(fullPath) && Path.GetPathRoot(fullPath) is { Length: > 0 } root
        && new DriveInfo(root).DriveType != DriveType.Network;

    /// <summary>Whether the file, or a folder above it, is a symbolic link or junction.</summary>
    /// <remarks>
    /// - OneDrive placeholders are reparse points too, but have no link target, so they still load.
    /// </remarks>
    private static bool GoesThroughLink(string fullPath)
    {
        var file = new FileInfo(fullPath);
        if (file.LinkTarget is not null)
        {
            return true;
        }

        for (var folder = file.Directory; folder is not null; folder = folder.Parent)
        {
            if (folder.LinkTarget is not null)
            {
                return true;
            }
        }

        return false;
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
