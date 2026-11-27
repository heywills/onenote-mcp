using System.ComponentModel;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

[McpServerToolType]
public sealed class OneNoteTools(OneNoteClient oneNote)
{
    private const int DefaultMaxImageKb = 512;

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    };

    // Navigation

    [McpServerTool(Name = "onenote_list_notebooks", ReadOnly = true)]
    [Description("List all open notebooks with their IDs, names, paths, and last modified times.")]
    public string ListNotebooks() =>
        ToJson(ParseNotebooks(oneNote.GetHierarchy("", HierarchyScope.hsNotebooks))
            .Select(nb => new { nb.Id, nb.Name, nb.Path, nb.LastModified }));

    [McpServerTool(Name = "onenote_list_sections", ReadOnly = true)]
    [Description("List all sections in a notebook, including sections inside section groups.")]
    public string ListSections([Description("The notebook ID, from onenote_list_notebooks")] string notebookId)
    {
        var notebook = ParseNotebooks(oneNote.GetHierarchy(notebookId, HierarchyScope.hsPages)).FirstOrDefault()
            ?? throw new McpException("Notebook not found");
        return ToJson(FlattenSections(notebook.Sections, notebook.SectionGroups, null));
    }

    [McpServerTool(Name = "onenote_list_pages", ReadOnly = true)]
    [Description("List all pages in a section with titles and last modified times.")]
    public string ListPages([Description("The section ID, from onenote_list_sections")] string sectionId)
    {
        var section = ParseNotebooks(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages))
            .SelectMany(nb => AllSections(nb.Sections, nb.SectionGroups))
            .FirstOrDefault(s => s.Id == sectionId)
            ?? throw new McpException("Section not found");
        return ToJson(section.Pages);
    }

    [McpServerTool(Name = "onenote_get_notebook_tree", ReadOnly = true)]
    [Description("Get the full hierarchy: notebooks, section groups, sections, and page titles.")]
    public string GetNotebookTree([Description("Notebook ID to scope the tree. Empty for all notebooks.")] string notebookId = "") =>
        ToJson(ParseNotebooks(oneNote.GetHierarchy(notebookId, HierarchyScope.hsPages)).Select(nb => new
        {
            nb.Id,
            nb.Name,
            Sections = nb.Sections.Select(SectionTree),
            SectionGroups = nb.SectionGroups.Select(GroupTree),
        }));

    // Content

    [McpServerTool(Name = "onenote_get_page", ReadOnly = true)]
    [Description("Get a page's content as Markdown. Images are listed as [Image N] with callback IDs for onenote_get_image.")]
    public string GetPage([Description("The page ID, from onenote_list_pages or a search")] string pageId)
    {
        var (markdown, images) = PageMarkdown.Convert(oneNote.GetPageContent(pageId));
        if (images.Count == 0)
        {
            return markdown;
        }

        var refs = images.Select(img =>
        {
            var dims = img is { Width: { } w, Height: { } h } ? $" ({w:0}x{h:0})" : "";
            return $"- [Image {img.Index}]: callback_id=`{img.CallbackId}`{dims}";
        });
        return $"{markdown}\n\n---\n**Image References:**\n{string.Join("\n", refs)}\n";
    }

    [McpServerTool(Name = "onenote_get_page_by_link", ReadOnly = true)]
    [Description("""
        Get a page as Markdown from its OneNote "Copy link to page" link.
        Accepts the onenote: link, the https://.../_layouts/Doc.aspx web link, or both pasted together.
        The notebook must be open in OneNote desktop. Output starts with the page ID for follow-up calls.
        """)]
    public string GetPageByLink([Description("The link, or both links, from \"Copy link to page\"")] string link)
    {
        var pageLink = PageLink.Parse(link) ?? throw new McpException("Not a OneNote page link");

        var sectionIds = XElement.Parse(oneNote.GetHierarchy("", HierarchyScope.hsSections))
            .Descendants(One + "Section")
            .Where(s => pageLink.MatchesSection((string?)s.Attribute("path") ?? ""))
            .Select(s => (string)s.Attribute("ID")!)
            .ToList();
        if (sectionIds.Count == 0)
        {
            throw new McpException("Section not found. Is the notebook open in OneNote desktop?");
        }

        var marker = $"page-id={{{pageLink.PageGuid}}}";
        foreach (var sectionId in sectionIds)
        {
            var pageId = XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages))
                .Descendants(One + "Page")
                .Select(p => (string)p.Attribute("ID")!)
                .FirstOrDefault(id => oneNote.GetHyperlink(id).Contains(marker, StringComparison.OrdinalIgnoreCase));
            if (pageId is not null)
            {
                return $"page_id: `{pageId}`\n\n{GetPage(pageId)}";
            }
        }

        throw new McpException($"Page not found in its section ({string.Join(", ", sectionIds)})");
    }

    [McpServerTool(Name = "onenote_get_page_raw", ReadOnly = true)]
    [Description("Get a page's raw OneNote XML, for debugging.")]
    public string GetPageRaw([Description("The page ID")] string pageId) => oneNote.GetPageContent(pageId);

    [McpServerTool(Name = "onenote_get_page_images", ReadOnly = true)]
    [Description("Get the images on a page, so you can see them.")]
    public IEnumerable<ContentBlock> GetPageImages(
        [Description("The page ID")] string pageId,
        [Description("Maximum number of images")] int maxImages = 10,
        [Description("Maximum size per image in KB. Larger images are shrunk.")] int maxSizeKb = DefaultMaxImageKb)
    {
        var (_, images) = PageMarkdown.Convert(oneNote.GetPageContent(pageId));
        if (images.Count == 0)
        {
            yield return Text("No images found on this page.");
            yield break;
        }

        foreach (var img in images.Take(maxImages))
        {
            ContentBlock block;
            try
            {
                block = Image(pageId, img.CallbackId, maxSizeKb);
            }
            catch (McpException e)
            {
                block = Text($"[Image {img.Index}] Error: {e.Message}");
            }

            if (block is ImageContentBlock)
            {
                yield return Text($"[Image {img.Index}] (callback_id: {img.CallbackId})");
            }

            yield return block;
        }
    }

    [McpServerTool(Name = "onenote_get_image", ReadOnly = true)]
    [Description("Get one image by its callback ID, so you can see it.")]
    public ContentBlock GetImage(
        [Description("The page ID")] string pageId,
        [Description("The image's callback ID, from onenote_get_page")] string callbackId,
        [Description("Maximum size in KB. Larger images are shrunk.")] int maxSizeKb = DefaultMaxImageKb) =>
        Image(pageId, callbackId, maxSizeKb);

    // Search

    [McpServerTool(Name = "onenote_search", ReadOnly = true)]
    [Description("Full-text search across all open notebooks. Uses Windows Search, so it finds indexed pages only.")]
    public string Search([Description("Search query")] string query) => SearchIn(query, "");

    [McpServerTool(Name = "onenote_search_in_notebook", ReadOnly = true)]
    [Description("Full-text search within one notebook. Uses Windows Search, so it finds indexed pages only.")]
    public string SearchInNotebook([Description("The notebook ID")] string notebookId, [Description("Search query")] string query) =>
        SearchIn(query, notebookId);

    // Write

    [McpServerTool(Name = "onenote_create_page", Destructive = false)]
    [Description("Create a new page in a section.")]
    public string CreatePage(
        [Description("The section ID")] string sectionId,
        [Description("Page title")] string title,
        [Description("""
            Optional HTML for the page body.
            Block: <p>, <div>, <br>, <h1>-<h6>, <ul>/<ol>/<li> (nested OK), <table>/<tr>/<th>/<td>, <pre>, <hr>.
            Inline: <b>, <i>, <u>, <s>, <sup>, <sub>, <code>, <a href>.
            """)] string contentHtml = "")
    {
        var pageId = oneNote.CreateNewPage(sectionId);
        try
        {
            var page = XElement.Parse(oneNote.GetPageContent(pageId));
            var titleText = page.Element(One + "Title")?.Descendants(One + "T").FirstOrDefault();
            titleText?.ReplaceNodes(new XCData(System.Net.WebUtility.HtmlEncode(title)));
            if (!string.IsNullOrWhiteSpace(contentHtml))
            {
                HtmlToOneNote.AddToPage(page, contentHtml);
            }

            oneNote.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting));
        }
        catch
        {
            // Don't leave a blank "Untitled page" behind
            try { oneNote.DeleteHierarchy(pageId); } catch (McpException) { }
            throw;
        }

        return ToJson(new { Status = "created", PageId = pageId, Title = title });
    }

    // Helpers

    private string SearchIn(string query, string notebookId)
    {
        var results = ParseSearchResults(oneNote.FindPages(query, notebookId));
        return results.Count > 0 ? ToJson(results) : ToJson(new { Message = "No results found", Query = query });
    }

    private ImageContentBlock Image(string pageId, string callbackId, int maxSizeKb)
    {
        var (data, mime) = Images.FitToSize(oneNote.GetBinaryPageContent(pageId, callbackId), maxSizeKb);
        return ImageContentBlock.FromBytes(data, mime);
    }

    private static TextContentBlock Text(string text) => new() { Text = text };

    private static string ToJson(object value) => JsonSerializer.Serialize(value, Json);

    private static IEnumerable<SectionInfo> AllSections(IEnumerable<SectionInfo> sections, IEnumerable<SectionGroupInfo> groups) =>
        sections.Concat(groups.SelectMany(g => AllSections(g.Sections, g.SectionGroups)));

    private static IEnumerable<object> FlattenSections(IEnumerable<SectionInfo> sections, IEnumerable<SectionGroupInfo> groups, string? prefix) =>
        sections.Select(s => (object)new { s.Id, s.Name, Group = prefix, s.Path, PageCount = s.Pages.Count })
            .Concat(groups.SelectMany(g => FlattenSections(g.Sections, g.SectionGroups, prefix is null ? g.Name : $"{prefix}/{g.Name}")));

    private static object SectionTree(SectionInfo s) =>
        new { s.Id, s.Name, Pages = s.Pages.Select(p => new { p.Id, p.Name, p.Level }) };

    private static object GroupTree(SectionGroupInfo g) =>
        new { g.Id, g.Name, Sections = g.Sections.Select(SectionTree), SectionGroups = g.SectionGroups.Select(GroupTree) };
}
