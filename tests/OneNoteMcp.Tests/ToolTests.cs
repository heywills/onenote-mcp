using System.Text.Json;
using Microsoft.Office.Interop.OneNote;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;

namespace OneNoteMcp.Tests;

/// <summary>
/// Tool behavior when OneNote or the file system fails, with <see cref="FakeOneNote"/> in place of OneNote.
/// </summary>
public sealed class ToolTests : IDisposable
{
    private static readonly byte[] Png = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private readonly FakeOneNote _oneNote = new();
    private readonly OneNoteTools _tools;
    private readonly string _folder = Path.Combine(Path.GetTempPath(), $"onenote-mcp-test-{Guid.NewGuid():N}");

    public ToolTests() => _tools = new(_oneNote);

    public void Dispose()
    {
        if (Directory.Exists(_folder))
        {
            Directory.Delete(_folder, recursive: true);
        }
    }

    [Fact]
    public void DeletesTheNewPageWhenTheUpdateFails()
    {
        _oneNote.NewPage = _ => "new";
        _oneNote.PageContent = Only("new", FakeOneNote.Page());
        _oneNote.Update = _ => throw new McpException("Update failed");

        Assert.Equal("Update failed", Assert.Throws<McpException>(() => _tools.CreatePage("section", "T", "<p>x</p>")).Message);
        Assert.Equal([("new", true)], _oneNote.Deleted);
    }

    [Fact]
    public void ReportsTheUpdateErrorWhenCleanupAlsoFails()
    {
        _oneNote.NewPage = _ => "new";
        _oneNote.PageContent = Only("new", FakeOneNote.Page());
        _oneNote.Update = _ => throw new McpException("Update failed");
        _oneNote.Delete = (_, _) => throw new InvalidOperationException("Delete failed");

        Assert.Equal("Update failed", Assert.Throws<McpException>(() => _tools.CreatePage("section", "T", "<p>x</p>")).Message);
        Assert.Equal([("new", true)], _oneNote.Deleted);
    }

    [Fact]
    public void CreatesATitleOnlyPage()
    {
        _oneNote.NewPage = _ => "new";
        _oneNote.PageContent = Only("new", FakeOneNote.Page());

        Assert.Contains("\"page_id\": \"new\"", _tools.CreatePage("section", "Q&A"));
        var page = System.Xml.Linq.XElement.Parse(Assert.Single(_oneNote.Updated));
        Assert.Equal("Q&amp;A", page.Element(System.Xml.Linq.XName.Get("Title", FakeOneNote.Ns))!.Value);
        Assert.Empty(_oneNote.Deleted);
    }

    [Fact]
    public void SavesPagesWithImagesItCantRead()
    {
        _oneNote.PageContent = Only("page", FakeOneNote.Page(Image("cb1", "Rental\nOptions") + Image("cb2")));
        _oneNote.Image = ImagesOf("page", id => id == "cb1" ? throw new McpException("OneNote returned no image data.") : Png);

        var saved = Save("page.md");

        Assert.Equal((1, 1), (saved.GetProperty("images").GetInt32(), saved.GetProperty("missing_images").GetInt32()));
        var markdown = File.ReadAllText(Path.Combine(_folder, "page.md"));
        Assert.Contains("[Image 1: OneNote returned no image data. Text in image: \"Rental Options\"]", markdown);
        Assert.Contains("page_images/image-2.png", markdown);
        Assert.True(File.Exists(Path.Combine(_folder, "page_images", "image-2.png")));
    }

    [Fact]
    public void SavesPagesWhenAnImageCantBeWritten()
    {
        _oneNote.PageContent = Only("page", FakeOneNote.Page(Image("cb1", "Q&A")));
        _oneNote.Image = ImagesOf("page", _ => Png);
        // A file where the image folder goes
        Directory.CreateDirectory(_folder);
        File.WriteAllText(Path.Combine(_folder, "page_images"), "");

        var saved = Save("page.html", format: "html");

        Assert.Equal((0, 1), (saved.GetProperty("images").GetInt32(), saved.GetProperty("missing_images").GetInt32()));
        var html = File.ReadAllText(Path.Combine(_folder, "page.html"));
        Assert.Contains("[Image 1: Can&#39;t save", html);
        Assert.Contains("Text in image: &quot;Q&amp;A&quot;]", html);
    }

    [Fact]
    public void SearchesTextWhenThePagesCantBeListed()
    {
        _oneNote.Hierarchy = (_, _) => throw new McpException("Notebook offline");
        _oneNote.Find = (_, _) => Pages(("text", "Notes"));

        var found = Search("project");

        Assert.Equal((false, true), (found.GetProperty("title_search").GetBoolean(), found.GetProperty("text_search").GetBoolean()));
        Assert.Equal(["text"], Ids(found));
    }

    [Fact]
    public void SearchesTitlesWhenWindowsSearchFails()
    {
        _oneNote.Hierarchy = (_, _) => Pages(("title", "Project plan"), ("other", "Notes"));
        _oneNote.Find = (_, _) => throw new McpException("Windows Search is off");

        var found = Search("project");

        Assert.Equal((true, false), (found.GetProperty("title_search").GetBoolean(), found.GetProperty("text_search").GetBoolean()));
        Assert.Equal(["title"], Ids(found));
    }

    [Fact]
    public void FailsWhenNoSearchWorks()
    {
        _oneNote.Hierarchy = (_, _) => throw new McpException("Notebook offline");
        _oneNote.Find = (_, _) => throw new McpException("Windows Search is off");

        Assert.Equal("Windows Search is off", Assert.Throws<McpException>(() => _tools.Search("project")).Message);
    }

    /// <summary>Text search only, so there's no page listing and nothing to fall back on.</summary>
    [Fact]
    public void FailsWhenWindowsSearchFailsForAnOperatorQuery()
    {
        _oneNote.Find = (_, _) => throw new McpException("Windows Search is off");

        Assert.Equal("Windows Search is off", Assert.Throws<McpException>(() => _tools.Search("project OR plan")).Message);
    }

    [Fact]
    public void FindsALinkedPageInALaterSection()
    {
        var looked = new List<string>();
        _oneNote.Hierarchy = (id, scope) => scope == HierarchyScope.hsSections ? Sections("s1", "s2") : Pages((id + "-page", "P"));
        _oneNote.Hyperlink = id =>
        {
            looked.Add(id);
            return PageLink(id == "s2-page" ? Target : Other);
        };
        _oneNote.PageContent = Only("s2-page", FakeOneNote.Page());

        Assert.StartsWith("page_id: `s2-page`", _tools.GetPageByLink(Link, inline: true));
        Assert.Equal(["s1-page", "s2-page"], looked);
    }

    [Fact]
    public void ReportsALinkedPageItCantFind()
    {
        _oneNote.Hierarchy = (id, scope) => scope == HierarchyScope.hsSections ? Sections("s1", "s2") : Pages((id + "-page", "P"));
        var looked = new List<string>();
        _oneNote.Hyperlink = id =>
        {
            looked.Add(id);
            return PageLink(Other);
        };

        Assert.Contains("Page not found in its section (s1, s2)", Assert.Throws<McpException>(() => _tools.GetPageByLink(Link, inline: true)).Message);
        Assert.Equal(["s1-page", "s2-page"], looked);
    }

    [Fact]
    public void ListsImageCallbackIdsInline()
    {
        _oneNote.PageContent = Only("page", FakeOneNote.Page(Image("cb1")));

        Assert.EndsWith("**Image References:**\n- [Image 1]: callback_id=`cb1`\n", _tools.GetPage("page", inline: true));
    }

    [Fact]
    public void ReturnsTheImagesItCanRead()
    {
        _oneNote.PageContent = Only("page", FakeOneNote.Page(Image("cb1", "Demo SI") + Image("cb2")));
        _oneNote.Image = ImagesOf("page", id => id == "cb1" ? throw new McpException("Not downloaded") : PngImage());

        var blocks = _tools.GetPageImages("page").ToList();

        Assert.Collection(blocks,
            b => Assert.Equal("[Image 1] Error: Not downloaded. Text in image: \"Demo SI\"", Assert.IsType<TextContentBlock>(b).Text),
            b => Assert.Equal("[Image 2] (callback_id: cb2)", Assert.IsType<TextContentBlock>(b).Text),
            b => Assert.IsType<ImageContentBlock>(b));
    }

    [Theory]
    [InlineData("""<one:Section xmlns:one="{0}" ID="x" name="S"/>""", "Not a page: Section")]
    [InlineData("""<one:Page xmlns:one="{0}" ID="x" name="P" isInRecycleBin="true"/>""", "already in Deleted Pages")]
    public void DeletesPagesOnly(string self, string error)
    {
        _oneNote.Hierarchy = (_, _) => string.Format(self, FakeOneNote.Ns);

        Assert.Contains(error, Assert.Throws<McpException>(() => _tools.DeletePage("x")).Message);
        Assert.Empty(_oneNote.Deleted);
    }

    [Fact]
    public void MovesPagesToDeletedPages()
    {
        _oneNote.Hierarchy = (_, scope) => scope == HierarchyScope.hsSelf
            ? $"""<one:Page xmlns:one="{FakeOneNote.Ns}" ID="x" name="P"/>""" : throw new McpException("Unexpected scope");

        Assert.Contains("moved to Deleted Pages", _tools.DeletePage("x"));
        Assert.Equal([("x", false)], _oneNote.Deleted);
    }

    private const string SectionPath = "https://x.test/Notes/S.one";
    private const string Target = "11111111-2222-3333-4444-555555555555";
    private const string Other = "66666666-7777-8888-9999-000000000000";
    private static readonly string Link = PageLink(Target);

    private static string PageLink(string pageGuid) => $"onenote:{SectionPath}#P&section-id={{{Other}}}&page-id={{{pageGuid}}}&end";

    private static string Sections(params string[] ids) => $"""
        <one:Notebooks xmlns:one="{FakeOneNote.Ns}"><one:Notebook name="N" ID="nb">
        {string.Concat(ids.Select(id => $"""<one:Section name="S" ID="{id}" path="{SectionPath}"/>"""))}
        </one:Notebook></one:Notebooks>
        """;

    /// <summary>A real 1x1 PNG, which the image tools can decode.</summary>
    private static byte[] PngImage()
    {
        using var bitmap = new System.Drawing.Bitmap(1, 1);
        using var stream = new MemoryStream();
        bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
        return stream.ToArray();
    }

    /// <summary>Returns <paramref name="xml"/> for <paramref name="pageId"/>, and fails the test for any other page.</summary>
    private static Func<string, string> Only(string pageId, string xml) =>
        id => id == pageId ? xml : throw new InvalidOperationException($"Read page {id}, expected {pageId}");

    /// <summary>The images of <paramref name="pageId"/> by callback ID. Fails the test for any other page.</summary>
    private static Func<string, string, byte[]> ImagesOf(string pageId, Func<string, byte[]> image) =>
        (id, callbackId) => id == pageId ? image(callbackId) : throw new InvalidOperationException($"Read an image of page {id}, expected {pageId}");

    /// <param name="text">The text OneNote recognized in the image</param>
    private static string Image(string callbackId, string? text = null) =>
        $"""<one:OE><one:Image>{(text is null ? "" : $"<one:OCRData lang=\"en-US\"><one:OCRText><![CDATA[{text}]]></one:OCRText></one:OCRData>")}<one:CallbackID callbackID="{callbackId}"/></one:Image></one:OE>""";

    private static string Pages(params (string Id, string Name)[] pages) => $"""
        <one:Notebooks xmlns:one="{FakeOneNote.Ns}"><one:Notebook name="N" ID="nb"><one:Section name="S" ID="s">
        {string.Concat(pages.Select(p => $"""<one:Page ID="{p.Id}" name="{p.Name}" lastModifiedTime="2026-01-01T00:00:00.000Z"/>"""))}
        </one:Section></one:Notebook></one:Notebooks>
        """;

    private JsonElement Save(string name, string format = "markdown") =>
        JsonDocument.Parse(_tools.GetPage("page", Path.Combine(_folder, name), format: format)).RootElement;

    private JsonElement Search(string query) => JsonDocument.Parse(_tools.Search(query)).RootElement;

    private static List<string?> Ids(JsonElement found) =>
        [.. found.GetProperty("results").EnumerateArray().Select(r => r.GetProperty("page_id").GetString())];
}
