using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;

namespace OneNoteMcp.Tests;

/// <summary>
/// Runs only when <c>ONENOTE_LIVE_SECTION</c> is set.
/// </summary>
public sealed class LiveFactAttribute : FactAttribute
{
    public LiveFactAttribute()
    {
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable(LiveTests.SectionVariable)))
        {
            Skip = $"Set {LiveTests.SectionVariable} to run tests against OneNote desktop";
        }
    }
}

/// <summary>
/// Tests against OneNote desktop. They create pages in a test section and delete them permanently.
/// </summary>
/// <remarks>
/// - <c>ONENOTE_LIVE_SECTION</c>: <c>Notebook name/Section name</c> for test pages
/// - <c>ONENOTE_LIVE_IMAGE_LINK</c>: optional "Copy link to page" link to a page with downloaded images
/// </remarks>
public class LiveTests
{
    public const string SectionVariable = "ONENOTE_LIVE_SECTION";

    private readonly OneNoteClient _client = new();
    private readonly OneNoteTools _tools;

    public LiveTests() => _tools = new OneNoteTools(_client);

    [LiveFact]
    public void RoundTripsMarkdownThroughOneNote()
    {
        var markdown = $"""
            # Live round trip test (deleting now)

            {MarkdownTests.Sample}

            > Quoted line

            3. Starts at three
            4. Four

            Literal \*stars\*, a # hash, snake_case, and 2 < 3 & 4

            \# Not a heading

            `npm install`

            ```
            first block
            ```

            ```
            second block
            ```

            > - quoted item
            > - another

            > First quoted paragraph with *a* and *b*
            >
            > Second quoted paragraph

            ---

            \>note and a ---

            \---

            [A link](https://x.test/a%20b)
            """.Replace("\r\n", "\n");

        // OneNote numbers each list itself, from 1
        var expected = markdown.Replace("3. Starts at three\n4. Four", "1. Starts at three\n2. Four");
        WithPage(markdown, pageId => Assert.Equal(expected, _tools.GetPage(pageId)));
    }

    [LiveFact]
    public void OpensPageByLink() =>
        WithPage("# Live link test (deleting now)\n\nBody", pageId =>
            Assert.StartsWith($"page_id: `{pageId}`", _tools.GetPageByLink(_client.GetHyperlink(pageId))));

    [LiveFact]
    public void ReportsOneNoteErrors()
    {
        var e = Assert.Throws<McpException>(() => _tools.GetPage("{bogus}"));
        Assert.Contains("Page not found (0x80042005)", e.Message);
    }

    [LiveFact]
    public void ReturnsImages()
    {
        var link = Environment.GetEnvironmentVariable("ONENOTE_LIVE_IMAGE_LINK");
        if (string.IsNullOrEmpty(link))
        {
            return; // optional
        }

        var pageId = _tools.GetPageByLink(link).Split('`')[1];
        var image = _tools.GetPageImages(pageId, maxImages: 1, maxSizeKb: 64).OfType<ImageContentBlock>().First();
        Assert.StartsWith("image/", image.MimeType);
        Assert.InRange(image.DecodedData.Length, 1, 64 * 1024);
    }

    private void WithPage(string markdown, Action<string> test)
    {
        var pageId = JsonDocument.Parse(_tools.CreatePageFromMarkdown(SectionId(), markdown)).RootElement.GetProperty("page_id").GetString()!;
        try
        {
            test(pageId);
        }
        finally
        {
            _client.DeleteHierarchy(pageId, permanently: true);
        }
    }

    private string SectionId()
    {
        var path = Environment.GetEnvironmentVariable(SectionVariable)!;
        var slash = path.LastIndexOf('/');
        var (notebookName, sectionName) = (path[..slash], path[(slash + 1)..]);

        var notebookId = JsonDocument.Parse(_tools.ListNotebooks()).RootElement.EnumerateArray()
            .First(n => n.GetProperty("name").GetString() == notebookName).GetProperty("id").GetString()!;
        return JsonDocument.Parse(_tools.ListSections(notebookId)).RootElement.EnumerateArray()
            .First(s => s.GetProperty("name").GetString() == sectionName).GetProperty("id").GetString()!;
    }
}
