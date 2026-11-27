using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

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
                return (png, "image/png");
            }
        }

        // Last resort: 400 px thumbnail
        var ratio = Math.Min(1.0, Math.Min(400.0 / source.Width, 400.0 / source.Height));
        return (ResizeToPng(source, Math.Max(1, (int)(source.Width * ratio)), Math.Max(1, (int)(source.Height * ratio))), "image/png");
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
