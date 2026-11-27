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
    public void NeedsANameBeforeTheExtension() =>
        Assert.Contains("name before the extension", Assert.Throws<McpException>(() => _tools.GetPage("{page}", @"C:\work\.md")).Message);

    [Fact]
    public void RejectsPathAndInlineTogether() =>
        Assert.Contains("not both", Assert.Throws<McpException>(() => _tools.GetPage("{page}", @"C:\work\page.md", inline: true)).Message);
}
