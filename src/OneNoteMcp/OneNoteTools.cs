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
    private const string OutputPathDescription = "Absolute path of the file to save, such as C:\\work\\page.md. Overwrites it. Required unless inline is true.";
    private const string GetPageDescription = """
        Save a page as Markdown to outputPath, with its images in a folder next to it, linked from the Markdown: page.md gets page_images.
        Or return it with inline=true, where images are listed as [Image N] with callback IDs for onenote_get_image.
        """;

    private const string ImagesDescription = """

        Local PNG, JPEG, GIF, BMP, and TIFF images are embedded. From inputPath: only images inside its folder, by relative path. Inline content: absolute paths.
        Web URLs, network shares and drives, unsafe symbolic links, files over 50 MB, images that would push the page past 50 MB, images inside headings, links, or formatting, and missing or other files stay links. The result counts both.
        """;

    private const string InlineDescription = "Return the content in the response instead of saving a file. Don't pass outputPath with it.";

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

    [McpServerTool(Name = "onenote_get_page", ReadOnly = false, Destructive = true)]
    [Description(GetPageDescription)]
    public string GetPage(
        [Description("The page ID, from onenote_list_pages or a search")] string pageId,
        [Description(OutputPathDescription)] string outputPath = "",
        [Description(InlineDescription)] bool inline = false)
    {
        CheckPageOutput(outputPath, inline);
        return inline ? PageText(pageId) : SavePage(pageId, outputPath);
    }

    /// <summary>Saves the page as Markdown, with its images in a folder next to it: <c>page.md</c> gets <c>page_images</c>.</summary>
    /// <remarks>
    /// - An image that can't be saved stays as text with the reason, such as <c>[Image N: OneNote hasn't downloaded this image yet…]</c>, so one image doesn't fail the save.
    /// - Images are saved full size. Saving again overwrites them.
    /// </remarks>
    private string SavePage(string pageId, string outputPath)
    {
        var xml = oneNote.GetPageContent(pageId);
        var (_, images) = PageMarkdown.Convert(xml);
        var folder = Path.GetFileNameWithoutExtension(outputPath) + "_images";
        var saved = images.ToDictionary(img => img.CallbackId, img => SaveImage(pageId, img, outputPath, folder));
        var (markdown, _) = PageMarkdown.Convert(xml, img => saved[img.CallbackId].Markdown);

        return ToJson(new
        {
            Status = "saved",
            Path = PageFiles.Write(outputPath, markdown),
            PageId = pageId,
            Images = saved.Values.Count(s => s.Saved),
            MissingImages = saved.Values.Count(s => !s.Saved),
        });
    }

    /// <summary>Saves one image to the page's image folder.</summary>
    /// <returns>The image's Markdown, and whether its file was saved</returns>
    private (string Markdown, bool Saved) SaveImage(string pageId, ImageRef img, string outputPath, string folder)
    {
        byte[] data;
        try
        {
            data = oneNote.GetBinaryPageContent(pageId, img.CallbackId);
        }
        catch (McpException e)
        {
            return (PageMarkdown.ImageProblem(img.Index, e.Message), false);
        }

        var extension = Images.Extension(data);
        try
        {
            PageFiles.Write(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(outputPath))!, folder, $"image-{img.Index}.{extension}"), data);
        }
        catch (McpException e)
        {
            return (PageMarkdown.ImageProblem(img.Index, e.Message), false);
        }

        return (PageMarkdown.ImageLink(img.Index, folder, extension), true);
    }

    private string PageText(string pageId)
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

    [McpServerTool(Name = "onenote_get_page_by_link", ReadOnly = false, Destructive = true)]
    [Description("""
        Get a page as Markdown from its OneNote "Copy link to page" link.
        Accepts the onenote: link, the https://.../_layouts/Doc.aspx web link, or both pasted together.
        The notebook must be open in OneNote desktop. Saves to outputPath with its images, like onenote_get_page, or returns it with inline=true. Either way, the result has the page ID for follow-up calls.
        """)]
    public string GetPageByLink(
        [Description("The link, or both links, from \"Copy link to page\"")] string link,
        [Description(OutputPathDescription)] string outputPath = "",
        [Description(InlineDescription)] bool inline = false)
    {
        CheckPageOutput(outputPath, inline);
        var pageId = FindPageByLink(link);
        return inline ? $"page_id: `{pageId}`\n\n{PageText(pageId)}" : SavePage(pageId, outputPath);
    }

    private string FindPageByLink(string link)
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
                return pageId;
            }
        }

        throw new McpException($"Page not found in its section ({string.Join(", ", sectionIds)})");
    }

    [McpServerTool(Name = "onenote_get_page_raw", ReadOnly = false, Destructive = true)]
    [Description("Save a page's raw OneNote XML to outputPath, or return it with inline=true. For debugging.")]
    public string GetPageRaw(
        [Description("The page ID")] string pageId,
        [Description(OutputPathDescription)] string outputPath = "",
        [Description(InlineDescription)] bool inline = false)
    {
        CheckOutput(outputPath, inline);
        return Output(oneNote.GetPageContent(pageId), outputPath, inline, pageId);
    }

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
        [Description("The image's callback ID, from onenote_get_page with inline=true, or onenote_get_page_images")] string callbackId,
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
            Block: <p>, <div>, <br>, <h1>-<h6>, <ul>/<ol>/<li> (nested OK), <table>/<tr>/<th>/<td>, <pre>, <blockquote>, <hr>.
            Inline: <b>, <i>, <u>, <s>, <sup>, <sub>, <code>, <a href>, <img>.
            """ + ImagesDescription)] string contentHtml = "",
        [Description("Absolute path of an HTML file for the page body, instead of contentHtml")] string inputPath = "") =>
        CreatePageFromHtml(sectionId, title, Input(contentHtml, inputPath, nameof(contentHtml)), ImageFolder(inputPath));

    [McpServerTool(Name = "onenote_create_page_from_markdown", Destructive = false)]
    [Description("""
        Create a new page in a section from Markdown.
        Supports headings, paragraphs, nested bullet and numbered lists, pipe tables, fenced code, blockquotes, rules, bold, italic, strikethrough, inline code, links, and images.
        onenote_get_page reads the page back as Markdown, and a page it saved comes back with its images.
        """ + ImagesDescription)]
    public string CreatePageFromMarkdown(
        [Description("The section ID")] string sectionId,
        [Description("Page content as Markdown. Required unless inputPath is given.")] string markdown = "",
        [Description("Page title. Empty to use the Markdown's leading \"# \" heading, which is then left out of the body.")] string title = "",
        [Description("Absolute path of a Markdown file, such as one onenote_get_page saved, instead of markdown")] string inputPath = "")
    {
        markdown = Input(markdown, inputPath, nameof(markdown));
        if (string.IsNullOrWhiteSpace(markdown) && string.IsNullOrWhiteSpace(title))
        {
            throw new McpException("Pass markdown or inputPath");
        }

        if (string.IsNullOrWhiteSpace(title))
        {
            (var heading, markdown) = MarkdownToOneNote.SplitTitle(markdown);
            title = heading ?? throw new McpException("No title: pass a title or start the Markdown with a \"# \" heading");
        }

        return CreatePageFromHtml(sectionId, title, MarkdownToOneNote.ToHtml(markdown), ImageFolder(inputPath), fromMarkdown: true);
    }

    [McpServerTool(Name = "onenote_delete_page", Destructive = true, Idempotent = false)]
    [Description("Move a page to its notebook's Deleted Pages, where OneNote keeps it for 60 days. Its subpages stay in the section. Deletes pages only, never sections or notebooks.")]
    public string DeletePage([Description("The page ID")] string pageId)
    {
        var page = XElement.Parse(oneNote.GetHierarchy(pageId, HierarchyScope.hsSelf));
        if (page.Name != One + "Page")
        {
            throw new McpException($"Not a page: {page.Name.LocalName}");
        }

        // OneNote's own error for this is "The page is read-only"
        if ((bool?)page.Attribute("isInRecycleBin") == true)
        {
            throw new McpException("The page is already in Deleted Pages");
        }

        oneNote.DeleteHierarchy(pageId);
        // No page ID: the page gets a new one in Deleted Pages
        return ToJson(new { Status = "moved to Deleted Pages", Title = (string?)page.Attribute("name") });
    }

    // Helpers

    /// <summary>Checks the output options before any OneNote call.</summary>
    private static void CheckOutput(string outputPath, bool inline)
    {
        var hasPath = !string.IsNullOrWhiteSpace(outputPath);
        if (inline)
        {
            if (hasPath)
            {
                throw new McpException("Pass outputPath or inline=true, not both");
            }

            return;
        }

        if (!hasPath)
        {
            throw new McpException("Pass outputPath, an absolute file path such as C:\\work\\page.md, or pass inline=true");
        }

        PageFiles.Check(outputPath, nameof(outputPath));
    }

    /// <summary>Checks the output options of a Markdown tool, whose image folder takes its name from the file.</summary>
    private static void CheckPageOutput(string outputPath, bool inline)
    {
        CheckOutput(outputPath, inline);
        if (!inline && Path.GetFileNameWithoutExtension(outputPath).Length == 0)
        {
            throw new McpException($"outputPath needs a name before the extension, such as page.md, not \"{Path.GetFileName(outputPath)}\"");
        }
    }

    /// <summary>The content, or the text of <paramref name="inputPath"/> when one is given.</summary>
    private static string Input(string content, string inputPath, string contentName)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            return content;
        }

        if (!string.IsNullOrWhiteSpace(content))
        {
            throw new McpException($"Pass {contentName} or inputPath, not both");
        }

        PageFiles.Check(inputPath, nameof(inputPath));
        return PageFiles.Read(inputPath);
    }

    private static string Output(string content, string outputPath, bool inline, string pageId) =>
        inline ? content : ToJson(new { Status = "saved", Path = PageFiles.Write(outputPath, content), PageId = pageId });

    /// <summary>The folder relative image paths resolve against: the input file's, or null for inline content.</summary>
    private static string? ImageFolder(string inputPath) =>
        string.IsNullOrWhiteSpace(inputPath) ? null : Path.GetDirectoryName(Path.GetFullPath(inputPath));

    /// <param name="imageFolder">The folder relative image paths resolve against. Null for inline content.</param>
    private string CreatePageFromHtml(string sectionId, string title, string contentHtml, string? imageFolder, bool fromMarkdown = false)
    {
        var pageId = oneNote.CreateNewPage(sectionId);
        var images = (Embedded: 0, Linked: 0);
        try
        {
            var page = XElement.Parse(oneNote.GetPageContent(pageId));
            var titleText = page.Element(One + "Title")?.Descendants(One + "T").FirstOrDefault();
            titleText?.ReplaceNodes(new XCData(System.Net.WebUtility.HtmlEncode(title)));
            if (!string.IsNullOrWhiteSpace(contentHtml))
            {
                images = HtmlToOneNote.AddToPage(page, contentHtml, fromMarkdown, (src, maxBytes) => PageFiles.ReadImage(src, imageFolder, maxBytes));
            }

            oneNote.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting));
        }
        catch
        {
            // Don't leave a blank "Untitled page" behind
            try { oneNote.DeleteHierarchy(pageId, permanently: true); } catch (McpException) { }
            throw;
        }

        return ToJson(new { Status = "created", PageId = pageId, Title = title, Images = images.Embedded, LinkedImages = images.Linked });
    }

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
