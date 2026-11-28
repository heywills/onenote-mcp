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
/// Tests against OneNote desktop. They create pages in a test section and delete them.
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
        WithPage(markdown, pageId => Assert.Equal(expected, _tools.GetPage(pageId, inline: true)));
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
        {
            Assert.StartsWith($"page_id: `{pageId}`", _tools.GetPageByLink(_client.GetHyperlink(pageId), inline: true));
            Assert.StartsWith($"<!DOCTYPE html>\n<!-- page_id: {pageId} -->\n<html>", _tools.GetPageByLink(_client.GetHyperlink(pageId), inline: true, format: "html"));
        });

    [LiveFact]
    public void SavesPagesToFiles() =>
        WithPage("# Live file test (deleting now)\n\nBody", pageId =>
        {
            var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
            var raw = Path.Combine(folder, "page.xml");
            try
            {
                var page = Path.Combine(folder, "new", "page.md");
                var saved = JsonDocument.Parse(_tools.GetPage(pageId, page)).RootElement;
                Assert.Equal(page, saved.GetProperty("path").GetString());
                Assert.Equal(pageId, saved.GetProperty("page_id").GetString());
                Assert.Equal(_tools.GetPage(pageId, inline: true), File.ReadAllText(page));

                var byLink = Path.Combine(folder, "by-link.md");
                _tools.GetPageByLink(_client.GetHyperlink(pageId), byLink);
                Assert.Equal(File.ReadAllText(page), File.ReadAllText(byLink));

                _tools.GetPageRaw(pageId, raw);
                Assert.Contains("Live file test", File.ReadAllText(raw));

                File.SetAttributes(raw, FileAttributes.ReadOnly);
                Assert.Contains("Can't save", Assert.Throws<McpException>(() => _tools.GetPageRaw(pageId, raw)).Message);
            }
            finally
            {
                if (File.Exists(raw))
                {
                    File.SetAttributes(raw, FileAttributes.Normal);
                }

                if (Directory.Exists(folder))
                {
                    Directory.Delete(folder, recursive: true);
                }
            }
        });

    [LiveFact]
    public void CreatesATitleOnlyPageFromMarkdown() =>
        WithCreated(_tools.CreatePageFromMarkdown(SectionId(), title: "Live title-only test (deleting now)"), pageId =>
            Assert.Equal("# Live title-only test (deleting now)", _tools.GetPage(pageId, inline: true)));

    // Windows Search hasn't indexed a new page, so only title search finds it
    [LiveFact]
    public void FindsNewPagesByTitleAndSkipsDeletedPages()
    {
        var token = $"zq{Guid.NewGuid():N}"[..12];
        WithCreated(_tools.CreatePageFromMarkdown(SectionId(), title: $"Live search {token} (deleting now)"), pageId =>
        {
            var found = JsonDocument.Parse(_tools.Search(token.ToUpperInvariant())).RootElement;
            var result = Assert.Single(found.GetProperty("results").EnumerateArray());
            Assert.Equal((pageId, true), (result.GetProperty("page_id").GetString(), result.GetProperty("in_title").GetBoolean()));
        });

        var deleting = JsonDocument.Parse(_tools.Search("deleting now", maxResults: 1000)).RootElement;
        Assert.DoesNotContain(deleting.GetProperty("results").EnumerateArray(), r => r.GetProperty("section").GetString() == "Deleted Pages");
    }

    // A page saved to a file comes back as the same Markdown when created from that file
    [LiveFact]
    public void CreatesPagesFromFiles() =>
        WithPage($"# Live input test (deleting now)\n\n{MarkdownTests.Sample}", pageId =>
        {
            var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
            try
            {
                var markdown = Path.Combine(folder, "page.md");
                _tools.GetPage(pageId, markdown);
                WithCreated(_tools.CreatePageFromMarkdown(SectionId(), inputPath: markdown), copyId =>
                    Assert.Equal(_tools.GetPage(pageId, inline: true), _tools.GetPage(copyId, inline: true)));

                var html = Path.Combine(folder, "page.html");
                File.WriteAllText(html, "<p><b>Bold</b> from a file</p>");
                WithCreated(_tools.CreatePage(SectionId(), "Live HTML input test (deleting now)", inputPath: html), copyId =>
                    Assert.Contains("**Bold** from a file", _tools.GetPage(copyId, inline: true)));
            }
            finally
            {
                if (Directory.Exists(folder))
                {
                    Directory.Delete(folder, recursive: true);
                }
            }
        });

    // OneNote stores a page link in its own relative form; the reader turns it back into one get_page_by_link opens.
    // The title has & # % : "Copy link to page" drops & and #, and the link still opens.
    [LiveFact]
    public void FollowsAPageLinkReadFromAPage() =>
        WithPage("# Live link target Q&A #1 50% (deleting now)\n\nTarget", targetId =>
            WithPage($"# Live link source (deleting now)\n\n[Target](<{_client.GetHyperlink(targetId)}>)", sourceId =>
            {
                var line = _tools.GetPage(sourceId, inline: true).Split('\n').Single(l => l.StartsWith("[Target]("));
                var link = line["[Target](<".Length..^2];
                Assert.StartsWith($"page_id: `{targetId}`", _tools.GetPageByLink(link, inline: true));
            }));

    /// <remarks>
    /// - Leaves the page in Deleted Pages: OneNote refuses to delete from there through the API
    /// </remarks>
    [LiveFact]
    public void DeletesPagesOnly()
    {
        var (notebookId, sectionId) = TestIds();
        Assert.Contains("Not a page: Section", Assert.Throws<McpException>(() => _tools.DeletePage(sectionId)).Message);

        var title = $"Live delete test {Guid.NewGuid():N} (in Deleted Pages)";
        var pageId = JsonDocument.Parse(_tools.CreatePageFromMarkdown(sectionId, $"# {title}\n\nx"))
            .RootElement.GetProperty("page_id").GetString()!;
        try
        {
            Assert.Contains("moved to Deleted Pages", _tools.DeletePage(pageId));
        }
        catch
        {
            try { _client.DeleteHierarchy(pageId, permanently: true); } catch (McpException) { }
            throw;
        }

        Assert.DoesNotContain(pageId, _tools.ListPages(sectionId));

        var deleted = System.Xml.Linq.XElement.Parse(_client.GetHierarchy(notebookId, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages))
            .Descendants(Hierarchy.One + "Page")
            .Single(p => (bool?)p.Attribute("isInRecycleBin") == true && (string?)p.Attribute("name") == title);
        Assert.Contains("already in Deleted Pages", Assert.Throws<McpException>(() => _tools.DeletePage((string)deleted.Attribute("ID")!)).Message);
    }

    [LiveFact]
    public void ReportsOneNoteErrors()
    {
        var e = Assert.Throws<McpException>(() => _tools.GetPage("{bogus}", inline: true));
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

        var pageId = _tools.GetPageByLink(link, inline: true).Split('`')[1];
        var image = _tools.GetPageImages(pageId, maxImages: 1, maxSizeKb: 64).OfType<ImageContentBlock>().First();
        Assert.StartsWith("image/", image.MimeType);
        Assert.InRange(image.DecodedData.Length, 1, 64 * 1024);
    }

    [LiveFact]
    public void SavesImagesNextToThePage()
    {
        var link = Environment.GetEnvironmentVariable("ONENOTE_LIVE_IMAGE_LINK");
        if (string.IsNullOrEmpty(link))
        {
            return; // optional
        }

        var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
        try
        {
            // '#' and spaces in the name: the link must still resolve
            var page = Path.Combine(folder, "Sprint #4.md");
            var saved = JsonDocument.Parse(_tools.GetPageByLink(link, page)).RootElement;
            Assert.True(saved.GetProperty("images").GetInt32() > 0);

            var markdown = File.ReadAllText(page);
            Assert.Contains("![Image 1](Sprint%20%234_images/image-1.", markdown);
            Assert.DoesNotContain("Image References", markdown);

            var image = Assert.Single(Directory.GetFiles(Path.Combine(folder, "Sprint #4_images"), "image-1.*"));
            Assert.StartsWith("image/", Images.MimeType(File.ReadAllBytes(image)));

            // Created again from the saved file, the page has the same images in the same places
            static string Body(string inline) => inline.Split("\n\n---\n**Image References:**")[0];
            var created = _tools.CreatePageFromMarkdown(SectionId(), inputPath: page);
            WithCreated(created, copyId =>
            {
                Assert.Equal(saved.GetProperty("images").GetInt32(), JsonDocument.Parse(created).RootElement.GetProperty("images").GetInt32());
                Assert.Equal(Body(_tools.GetPage(saved.GetProperty("page_id").GetString()!, inline: true)), Body(_tools.GetPage(copyId, inline: true)));
            });
        }
        finally
        {
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    // A page created from onenote_get_page's HTML reads back as the same HTML
    [LiveFact]
    public void RoundTripsHtmlThroughOneNote()
    {
        const string html = """
            <p style="font-family:Calibri;font-size:12.0pt;color:#14324D;text-align:center">Centered</p>
            <p>Text with <span style="color:#FF0000">red</span>, <span style="background:yellow">highlight</span>, and a<br>line break</p>
            <p>Parent</p>
            <div class="indent"><p>Indented child</p></div>
            <p></p>
            <h2>Heading</h2>
            <ul><li>Bullet<ol><li>Number</li></ol></li></ul>
            <table><tr><th>Head</th></tr><tr><td style="background-color:#DDEBF7">Shaded</td></tr></table>
            <pre>A
            B</pre><div class="indent"><p>Under B</p></div><pre>C</pre>
            """;
        WithCreated(_tools.CreatePage(SectionId(), "Live HTML round trip (deleting now)", html), pageId =>
        {
            var first = _tools.GetPage(pageId, inline: true, format: "html");
            // One code block, split to show the indent, with no blank line added
            Assert.Contains("</div>\n<pre>\nC</pre>", first);
            // OneNote writes #FF0000 as red, and copies the alignment into the style
            Assert.Contains("text-align:center\"", first);
            Assert.DoesNotContain("text-align:center;text-align", first);
            Assert.Contains("color:red", first);
            Assert.Contains("<br", first);
            Assert.Contains("class=\"indent\"", first);
            Assert.Contains("<th>", first);
            Assert.Contains("background-color:#DDEBF7", first);

            // OneNote line-wraps its own markup differently each time
            static string Collapse(string text) => System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
            WithCreated(_tools.CreatePage(SectionId(), contentHtml: first), copyId =>
                Assert.Equal(Collapse(first), Collapse(_tools.GetPage(copyId, inline: true, format: "html"))));
        });
    }

    [LiveFact]
    public void SavesHtmlWithImagesAndCreatesItAgain()
    {
        var link = Environment.GetEnvironmentVariable("ONENOTE_LIVE_IMAGE_LINK");
        if (string.IsNullOrEmpty(link))
        {
            return; // optional
        }

        var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
        try
        {
            var page = Path.Combine(folder, "page.html");
            var saved = JsonDocument.Parse(_tools.GetPageByLink(link, page, format: "html")).RootElement;
            Assert.Contains("<img src=\"page_images/image-1.", File.ReadAllText(page));

            var created = _tools.CreatePage(SectionId(), inputPath: page);
            WithCreated(created, copyId =>
                Assert.Equal(saved.GetProperty("images").GetInt32(), JsonDocument.Parse(created).RootElement.GetProperty("images").GetInt32()));
        }
        finally
        {
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    [LiveFact]
    public void SavesNoImageFolderForPagesWithoutImages() =>
        WithPage("# Live no-images test (deleting now)\n\nBody", pageId =>
        {
            var folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");
            try
            {
                var saved = JsonDocument.Parse(_tools.GetPage(pageId, Path.Combine(folder, "page.md"))).RootElement;
                Assert.Equal(0, saved.GetProperty("images").GetInt32());
                Assert.False(Directory.Exists(Path.Combine(folder, "page_images")));
            }
            finally
            {
                if (Directory.Exists(folder))
                {
                    Directory.Delete(folder, recursive: true);
                }
            }
        });

    private void WithPage(string markdown, Action<string> test) => WithCreated(_tools.CreatePageFromMarkdown(SectionId(), markdown), test);

    /// <param name="created">A create tool's result</param>
    private void WithCreated(string created, Action<string> test)
    {
        var pageId = JsonDocument.Parse(created).RootElement.GetProperty("page_id").GetString()!;
        try
        {
            test(pageId);
        }
        finally
        {
            _client.DeleteHierarchy(pageId, permanently: true);
        }
    }

    private string SectionId() => TestIds().Section;

    private (string Notebook, string Section) TestIds()
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
        return (notebookId, IdOf(_tools.ListSections(notebookId), sectionName)
            ?? throw new InvalidOperationException($"Section \"{sectionName}\" not found in \"{notebookName}\""));
    }
}
