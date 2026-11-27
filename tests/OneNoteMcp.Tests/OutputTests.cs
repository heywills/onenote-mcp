using System.Text;
using ModelContextProtocol;

namespace OneNoteMcp.Tests;

/// <summary>
/// Output path checks. They run before any OneNote call, so these tests don't need OneNote.
/// </summary>
public class OutputTests
{
    private readonly OneNoteTools _tools = new(new OneNoteClient());

    public static TheoryData<Func<OneNoteTools, string, string>> Tools => new()
    {
        (tools, path) => tools.GetPage("{page}", path),
        (tools, path) => tools.GetPageByLink("onenote:https://x.test/S.one#P&section-id={A}&page-id={B}&end", path),
        (tools, path) => tools.GetPageRaw("{page}", path),
    };

    [Theory]
    [MemberData(nameof(Tools))]
    public void NeedsAPathOrInline(Func<OneNoteTools, string, string> tool) =>
        Assert.Contains("Pass outputPath", Assert.Throws<McpException>(() => tool(_tools, "")).Message);

    [Theory]
    [MemberData(nameof(Tools))]
    public void NeedsAnAbsolutePath(Func<OneNoteTools, string, string> tool) =>
        Assert.Contains("must be absolute", Assert.Throws<McpException>(() => tool(_tools, @"notes\page.md")).Message);

    [Theory]
    [InlineData(@"C:\work\")]
    [InlineData(@"C:\")]
    [InlineData(@"\\server\share")]
    public void NeedsAFileNotAFolder(string path) =>
        Assert.Contains("must name a file", Assert.Throws<McpException>(() => _tools.GetPage("{page}", path)).Message);

    [Fact]
    public void RejectsAnExistingFolder() =>
        Assert.Contains("must name a file", Assert.Throws<McpException>(() => _tools.GetPage("{page}", Path.GetTempPath().TrimEnd('\\'))).Message);

    [Theory]
    [InlineData(@"C:\work\What now?.md")]
    [InlineData(@"C:\work\a|b.md")]
    [InlineData(@"C:\work\page.md:stream")]
    public void RejectsInvalidFileNames(string path) =>
        Assert.Contains("can't hold", Assert.Throws<McpException>(() => _tools.GetPage("{page}", path)).Message);

    [Theory]
    [InlineData(@"\\.\pipe\name")]
    [InlineData(@"\\?\C:\work\page.md")]
    [InlineData(@"C:\work\NUL")]
    public void RejectsDevicePaths(string path) =>
        Assert.Contains("device path", Assert.Throws<McpException>(() => _tools.GetPage("{page}", path)).Message);

    [Theory]
    [InlineData(@"C:\work\page.md.")]
    [InlineData(@"C:\work\page.md ")]
    [InlineData(@"C:\does-not-exist\.")]
    public void RejectsNamesWindowsRenames(string path) =>
        Assert.Contains("dot or space", Assert.Throws<McpException>(() => _tools.GetPage("{page}", path)).Message);

    [Theory]
    [InlineData(@"C:\work\CON.md")]
    [InlineData(@"C:\work\aux")]
    [InlineData(@"C:\work\COM9.txt")]
    public void RejectsReservedNames(string path) =>
        Assert.Contains("reserves", Assert.Throws<McpException>(() => _tools.GetPage("{page}", path)).Message);

    [Fact]
    public void RejectsContentAndInputPathTogether()
    {
        Assert.Contains("not both", Assert.Throws<McpException>(() => _tools.CreatePageFromMarkdown("{section}", "# T", inputPath: @"C:\work\page.md")).Message);
        Assert.Contains("not both", Assert.Throws<McpException>(() => _tools.CreatePage("{section}", "T", "<p>x</p>", @"C:\work\page.html")).Message);
    }

    [Fact]
    public void NeedsMarkdownOrInputPath() =>
        Assert.Contains("Pass markdown or inputPath", Assert.Throws<McpException>(() => _tools.CreatePageFromMarkdown("{section}")).Message);

    [Theory]
    [InlineData(@"notes\page.md", "must be absolute")]
    [InlineData(@"\\.\pipe\name", "device path")]
    [InlineData(@"\\?\C:\work\page.md", "device path")]
    [InlineData(@"C:\work\NUL", "device path")]
    [InlineData(@"C:\work\COM1.md", "reserves")]
    [InlineData(@"C:\work\CONIN$", "reserves")]
    [InlineData("C:\\work\\a\0b.md", "valid path")]
    public void ChecksTheInputPath(string path, string error) =>
        Assert.Contains(error, Assert.Throws<McpException>(() => _tools.CreatePageFromMarkdown("{section}", inputPath: path)).Message);

    [Fact]
    public void RejectsAFolderAsInput() =>
        Assert.Contains("must name a file", Assert.Throws<McpException>(() => _tools.CreatePage("{section}", "T", inputPath: Path.GetTempPath().TrimEnd('\\'))).Message);

    [Fact]
    public void TreatsBlankContentAsNone()
    {
        var path = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.md");
        Assert.Contains("Can't read", Assert.Throws<McpException>(() => _tools.CreatePageFromMarkdown("{section}", " ", inputPath: path)).Message);
    }

    [Theory]
    [InlineData("utf-8", true)]
    [InlineData("utf-8", false)]
    [InlineData("utf-16", true)]
    public void ReadsUnicodeFiles(string encoding, bool bom) =>
        WithFile(Encoding.GetEncoding(encoding), "# Café", path => Assert.Equal("# Café", PageFiles.Read(path)), bom);

    [Fact]
    public void RejectsFilesThatArentUtf8() =>
        WithFile(Encoding.Latin1, "# Café", path => Assert.Contains("isn't UTF-8", Assert.Throws<McpException>(() => PageFiles.Read(path)).Message));

    [Fact]
    public void RejectsUtf16WithoutAByteOrderMark() =>
        WithFile(Encoding.Unicode, "# Notes", path => Assert.Contains("null character", Assert.Throws<McpException>(() => PageFiles.Read(path)).Message));

    [Fact]
    public void RejectsEmptyFiles() =>
        WithFile(new UTF8Encoding(false), " \n", path => Assert.Contains("is empty", Assert.Throws<McpException>(() => PageFiles.Read(path)).Message));

    [Fact]
    public void ReadsLocalImages()
    {
        var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
        var image = Path.Combine(folder, "Sprint #4_images", "image-1.png");
        PageFiles.Write(image, [1, 2, 3]);
        try
        {
            Assert.Equal([1, 2, 3], PageFiles.ReadImage("Sprint%20%234_images/image-1.png", folder));
            Assert.Equal([1, 2, 3], PageFiles.ReadImage(image, null));
            Assert.Equal([1, 2, 3], PageFiles.ReadImage(new Uri(image).AbsoluteUri, null));

            // Inline content has no folder for relative paths
            Assert.Null(PageFiles.ReadImage("Sprint%20%234_images/image-1.png", null));
            Assert.Null(PageFiles.ReadImage("https://x.test/a.png", folder));
            Assert.Null(PageFiles.ReadImage("data:image/png;base64,AAAA", folder));
            Assert.Null(PageFiles.ReadImage("missing.png", folder));
            Assert.Null(PageFiles.ReadImage("CON", folder));

            // A file's images must sit inside its folder
            var inner = Path.Combine(folder, "inner");
            Directory.CreateDirectory(inner);
            Assert.Null(PageFiles.ReadImage("../Sprint%20%234_images/image-1.png", inner));
            Assert.Null(PageFiles.ReadImage("..%2FSprint%20%234_images%2Fimage-1.png", inner));
            Assert.Null(PageFiles.ReadImage(image, inner));
            Assert.Null(PageFiles.ReadImage(Uri.EscapeDataString(image), inner));
            Assert.Null(PageFiles.ReadImage(new Uri(image).AbsoluteUri, inner));
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    // Opening any of these would connect to the host and send the user's Windows credentials
    [Theory]
    [InlineData(@"\\host\share\a.png")]
    [InlineData("//host/share/a.png")]
    [InlineData("file://host/share/a.png")]
    [InlineData("file:////host/share/a.png")]
    [InlineData("%5C%5Chost%5Cshare%5Ca.png")]
    [InlineData("%2F%2Fhost%2Fshare%2Fa.png")]
    [InlineData(@"\/host/share/a.png")]
    [InlineData("file:///%2Fhost/share/a.png")]
    [InlineData("file:///%5Chost/share/a.png")]
    public void NeverReadsNetworkPaths(string src)
    {
        Assert.Null(PageFiles.ReadImage(src, null));
        Assert.Null(PageFiles.ReadImage(src, Path.GetTempPath()));
    }

    [Fact]
    public void ReadsImagesBesideAFileAtADriveRoot()
    {
        // The temp drive's root stands in for C:\: only the folder prefix logic is under test
        var root = Path.GetPathRoot(Path.GetTempPath())!;
        var name = $"onenote-mcp-test-{Guid.NewGuid():N}";
        var image = Path.Combine(Path.GetTempPath(), name, "a.png");
        PageFiles.Write(image, [1]);
        try
        {
            var relative = Path.GetRelativePath(root, image).Replace('\\', '/');
            Assert.Equal([1], PageFiles.ReadImage(relative, root));
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(image)!, recursive: true);
        }
    }

    [Fact]
    public void SkipsImagesThroughAJunction()
    {
        var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
        var outside = Path.Combine(folder, "outside");
        var page = Path.Combine(folder, "page");
        PageFiles.Write(Path.Combine(outside, "a.png"), [1]);
        Directory.CreateDirectory(page);

        // A junction needs no admin rights, unlike a symbolic link
        using (var mklink = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe", $"/c mklink /J \"{Path.Combine(page, "link")}\" \"{outside}\"") { CreateNoWindow = true, UseShellExecute = false })!)
        {
            mklink.WaitForExit();
        }

        try
        {
            Assert.True(Directory.Exists(Path.Combine(page, "link")));
            Assert.Null(PageFiles.ReadImage("link/a.png", page));
        }
        finally
        {
            // Remove the junction itself first, so the recursive delete can't follow it
            if (Directory.Exists(Path.Combine(page, "link")))
            {
                Directory.Delete(Path.Combine(page, "link"));
            }

            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void SkipsFilesOverTheSizeLimit()
    {
        var path = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}.png");
        using (var file = File.Create(path))
        {
            file.SetLength(PageFiles.MaxImageBytes + 1);
        }

        try
        {
            Assert.Null(PageFiles.ReadImage(path, null));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ShrinksLargeImagesForOneNote()
    {
        using var bitmap = new System.Drawing.Bitmap(3000, 3000);
        var random = new Random(1);
        for (var x = 0; x < 3000; x += 3)
        {
            for (var y = 0; y < 3000; y += 3)
            {
                bitmap.SetPixel(x, y, System.Drawing.Color.FromArgb(random.Next(256), random.Next(256), random.Next(256)));
            }
        }

        using var bmp = new MemoryStream();
        bitmap.Save(bmp, System.Drawing.Imaging.ImageFormat.Bmp);
        Assert.True(bmp.Length > Images.MaxEmbedKb * 1024);

        var image = Images.ForOneNote(bmp.ToArray());
        Assert.Equal("png", image?.Format);
        Assert.InRange(image!.Value.Data.Length, 1, Images.MaxEmbedKb * 1024);
    }

    [Fact]
    public void ConvertsImagesForOneNote()
    {
        using var bitmap = new System.Drawing.Bitmap(2, 2);
        using var gif = new MemoryStream();
        bitmap.Save(gif, System.Drawing.Imaging.ImageFormat.Gif);

        var converted = Images.ForOneNote(gif.ToArray());
        Assert.Equal("png", converted?.Format);
        Assert.Equal("image/png", Images.MimeType(converted!.Value.Data));
        Assert.Equal("jpg", Images.ForOneNote([0xFF, 0xD8, 0xFF, 0xE0])?.Format);
        Assert.Null(Images.ForOneNote("RIFF\0\0\0\0WEBP"u8.ToArray()));
        Assert.Null(Images.ForOneNote([0x47, 0x49, 0x46, 0x38, 0x00]));
    }

    private static void WithFile(Encoding encoding, string text, Action<string> test, bool bom = false)
    {
        var path = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}.md");
        File.WriteAllBytes(path, [.. bom ? encoding.GetPreamble() : [], .. encoding.GetBytes(text)]);
        try
        {
            test(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReportsAMissingInputFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.md");
        Assert.Contains("Can't read", Assert.Throws<McpException>(() => _tools.CreatePageFromMarkdown("{section}", inputPath: path)).Message);
    }

    [Fact]
    public void NeedsANameBeforeTheExtension()
    {
        Assert.Contains("name before the extension", Assert.Throws<McpException>(() => _tools.GetPage("{page}", @"C:\work\.md")).Message);
        Assert.Contains("name before the extension", Assert.Throws<McpException>(() => _tools.GetPageByLink("onenote:https://x.test/S.one#P&section-id={A}&page-id={B}&end", @"C:\work\.md")).Message);
    }

    [Fact]
    public void RejectsPathAndInlineTogether() =>
        Assert.Contains("not both", Assert.Throws<McpException>(() => _tools.GetPage("{page}", @"C:\work\page.md", inline: true)).Message);
}
