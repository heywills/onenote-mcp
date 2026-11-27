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
}
