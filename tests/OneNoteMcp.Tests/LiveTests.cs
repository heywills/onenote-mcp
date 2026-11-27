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

    private const string TabIndentedCode = "```\nfunc f() {\n\treturn 1\n}\n```";

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

            #### Four

            ###### Six

            ## Ticket \#

            - item

              > quote in the item

            - next

            ````
            ```
            nested fence
            ```
            ````

            [1\]: <https://example.com>

            - a

              > - b
              >   - c

            Use \&nbsp; but keep R&D, and see <https://x.test/_x_/a>

            [Spec](\\\\fileserver\\share\\spec.docx)

            ```

            leading blank line
            ```

            1. one
            2. two

            1) adjacent list

            - bullets

            * adjacent bullets

            `x  =  1` and > *whole quote italic* stays text

            > *whole quote italic*

            - before
            - ![screenshot](https://x.test/s.png)
            - after

            func<strong>(x)</strong> and [`a[0]`](https://x.test/code)

            - nested code

              ```
              a
              ```

            ```
            b
            ```

            - x
              - a

              * b

            > - quoted a
            >
            > * quoted b

            {TabIndentedCode}
            """.Replace("\r\n", "\n");

        // OneNote numbers each list itself, from 1
        // Images become links, and OneNote stores a \\server\share link as file://server/share
        var expected = markdown.Replace("3. Starts at three\n4. Four", "1. Starts at three\n2. Four").Replace("- ![screenshot]", "- [screenshot]")
            .Replace(@"(\\\\fileserver\\share\\spec.docx)", "(file://fileserver/share/spec.docx)")
            .Replace("\treturn 1", "        return 1"); // OneNote stores a tab in code as 8 spaces
        WithPage(markdown, pageId => Assert.Equal(expected, _tools.GetPage(pageId)));
    }

    // What OneNote shows, not what the reader prints: each list restarts at 1, in Arabic numerals.
    [LiveFact]
    public void RestartsEachNumberedListInOneNote() =>
        WithPage("# Live numbering test (deleting now)\n\n1. a\n2. b\n\nText\n\n1. c\n2. d\n\nText\n\n1. e\n   1. nested", pageId =>
            Assert.Equal(["1.", "2.", "1)", "2)", "1.", "1."], System.Xml.Linq.XElement.Parse(_client.GetPageContent(pageId))
                .Descendants(Hierarchy.One + "Number").Select(n => (string?)n.Attribute("text"))));

    [LiveFact]
    public void OpensPageByLink() =>
        WithPage("# Live link test (deleting now)\n\nBody", pageId =>
            Assert.StartsWith($"page_id: `{pageId}`", _tools.GetPageByLink(_client.GetHyperlink(pageId))));

    // OneNote stores a page link in its own relative form; the reader turns it back into one get_page_by_link opens.
    [LiveFact]
    public void FollowsAPageLinkReadFromAPage() =>
        WithPage("# Live link target (deleting now)\n\nTarget", targetId =>
            WithPage($"# Live link source (deleting now)\n\n[Target](<{_client.GetHyperlink(targetId)}>)", sourceId =>
            {
                var line = _tools.GetPage(sourceId).Split('\n').Single(l => l.StartsWith("[Target]("));
                var link = line["[Target](<".Length..^2];
                Assert.StartsWith($"page_id: `{targetId}`", _tools.GetPageByLink(link));
            }));

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
        if (slash <= 0)
        {
            throw new InvalidOperationException($"Set {SectionVariable} to \"Notebook name/Section name\", not \"{path}\"");
        }

        var (notebookName, sectionName) = (path[..slash], path[(slash + 1)..]);
        static string? IdOf(string json, string name) => JsonDocument.Parse(json).RootElement.EnumerateArray()
            .Where(e => e.GetProperty("name").GetString() == name).Select(e => e.GetProperty("id").GetString()).FirstOrDefault();

        var notebookId = IdOf(_tools.ListNotebooks(), notebookName)
            ?? throw new InvalidOperationException($"Notebook \"{notebookName}\" isn't open in OneNote");
        return IdOf(_tools.ListSections(notebookId), sectionName)
            ?? throw new InvalidOperationException($"Section \"{sectionName}\" not found in \"{notebookName}\"");
    }
}
