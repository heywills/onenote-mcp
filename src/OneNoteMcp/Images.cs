using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using ModelContextProtocol;

namespace OneNoteMcp;

/// <summary>
/// Detects image types and shrinks images over a size limit.
/// </summary>
public static class Images
{
    public static string MimeType(ReadOnlySpan<byte> data) => data switch
    {
        [0x89, 0x50, 0x4E, 0x47, ..] => "image/png",
        [0xFF, 0xD8, 0xFF, ..] => "image/jpeg",
        [0x47, 0x49, 0x46, 0x38, ..] => "image/gif",
        [0x42, 0x4D, ..] => "image/bmp",
        [0x49, 0x49, 0x2A, 0x00, ..] or [0x4D, 0x4D, 0x00, 0x2A, ..] => "image/tiff",
        [0x52, 0x49, 0x46, 0x46, _, _, _, _, 0x57, 0x45, 0x42, 0x50, ..] => "image/webp",
        _ => "application/octet-stream",
    };

    /// <summary>Larger images are shrunk before they go into a page, so the page update stays a size OneNote accepts.</summary>
    public const int MaxEmbedKb = 10 * 1024;

    /// <summary>The image as OneNote stores it: PNG or JPEG as is, other bitmaps as PNG, and anything over <see cref="MaxEmbedKb"/> shrunk.</summary>
    /// <returns>Null for data that isn't an image OneNote can hold, such as WebP or a vector format</returns>
    public static (byte[] Data, string Format)? ForOneNote(byte[] data)
    {
        var format = MimeType(data) switch
        {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" or "image/bmp" or "image/tiff" => "convert",
            _ => null,
        };

        try
        {
            return format switch
            {
                null => null,
                _ when data.Length > MaxEmbedKb * 1024 => (FitToSize(data, MaxEmbedKb).Data, "png"),
                // A small palette GIF can grow into a large PNG
                "convert" => (FitToSize(ToPng(data), MaxEmbedKb).Data, "png"),
                _ => (data, format),
            };
        }
        catch (Exception e) when (e is McpException or ArgumentException or ExternalException or OutOfMemoryException)
        {
            return null;
        }
    }

    private static byte[] ToPng(byte[] data)
    {
        using var source = Image.FromStream(new MemoryStream(data));
        return ResizeToPng(source, source.Width, source.Height);
    }

    /// <summary>The file extension for the image's type, without the dot</summary>
    public static string Extension(ReadOnlySpan<byte> data) => MimeType(data) switch
    {
        "image/jpeg" => "jpg",
        "image/tiff" => "tif",
        "application/octet-stream" => "bin",
        var mime => mime["image/".Length..],
    };

    /// <summary>
    /// Returns the image as is when it fits in <paramref name="maxKb"/>, or a smaller PNG.
    /// </summary>
    public static (byte[] Data, string MimeType) FitToSize(byte[] data, int maxKb)
    {
        var mime = MimeType(data);
        if (data.Length <= maxKb * 1024)
        {
            return (data, mime);
        }

        // GDI+ can't read every type OneNote stores, such as some vector formats
        try
        {
            return (ShrinkToPng(data, maxKb), "image/png");
        }
        catch (Exception e) when (e is ArgumentException or ExternalException or OutOfMemoryException)
        {
            throw new McpException($"Can't shrink this {mime} image below {maxKb} KB: {e.Message}", e);
        }
    }

    private static byte[] ShrinkToPng(byte[] data, int maxKb)
    {
        using var source = Image.FromStream(new MemoryStream(data));
        foreach (var scale in new[] { 0.75, 0.5, 0.35, 0.25, 0.15 })
        {
            var (w, h) = ((int)(source.Width * scale), (int)(source.Height * scale));
            if (w < 100 || h < 100)
            {
                break;
            }

            var png = ResizeToPng(source, w, h);
            if (png.Length <= maxKb * 1024)
            {
                return png;
            }
        }

        // Last resort: 400 px thumbnail
        var ratio = Math.Min(1.0, Math.Min(400.0 / source.Width, 400.0 / source.Height));
        return ResizeToPng(source, Math.Max(1, (int)(source.Width * ratio)), Math.Max(1, (int)(source.Height * ratio)));
    }

    private static byte[] ResizeToPng(Image source, int width, int height)
    {
        using var bitmap = new Bitmap(width, height);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(source, 0, 0, width, height);
        }

        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }
}
