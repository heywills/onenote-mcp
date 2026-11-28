using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using ModelContextProtocol;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

/// <summary>A moved page's old and new IDs.</summary>
public sealed record MovedPage(string OldId, string NewId, string Title);

/// <summary>The result of <see cref="PageMover.Move"/>.</summary>
/// <param name="PageId">The page's new ID</param>
/// <param name="Moved">Every moved page, the page and then its subpages</param>
public sealed record PageMove(string Status, string PageId, string Title, string Section, IReadOnlyList<MovedPage> Moved);

/// <summary>
/// Moves a page and its subpages with <c>UpdateHierarchy</c>, and finds their new IDs.
/// </summary>
/// <remarks>
/// - Reads only the page, its section, and the target, so an offline notebook elsewhere doesn't matter.
/// - A page that changes section gets a new ID. Its link's page-id stays, which is how the new ID is found.
/// </remarks>
public sealed partial class PageMover(IOneNoteClient oneNote)
{
    public PageMove Move(string pageId, string sectionId, string afterPageId)
    {
        var source = SourceSection(pageId);
        var page = source.Elements(One + "Page").First(p => PageMoves.SameId(PageMoves.Id(p), pageId));
        var target = TargetSection(sectionId);
        var moving = PageMoves.WithSubpages(page);
        var sameSection = PageMoves.SameId(PageMoves.Id(source), PageMoves.Id(target));

        var links = sameSection ? new Dictionary<string, XElement>() : MovingLinks(moving);
        var before = target.Elements(One + "Page").Select(PageMoves.Id).ToHashSet(StringComparer.OrdinalIgnoreCase);
        oneNote.UpdateHierarchy(PageMoves.Place(target, moving, afterPageId).ToString(SaveOptions.DisableFormatting));

        var moved = sameSection
            ? [.. moving.Select(p => new MovedPage(PageMoves.Id(p), PageMoves.Id(p), Name(p)))]
            : NewIds(sectionId, before, links);
        return new PageMove("moved", moved[0].NewId, Name(page), Name(target), moved);
    }

    /// <summary>The section holding the page, with its pages.</summary>
    /// <remarks>
    /// - A page ID starts with its section's: <c>{section}{1}{E…}</c> is in <c>{section}{1}{B0}</c>. The whole hierarchy is the fallback.
    /// </remarks>
    private XElement SourceSection(string pageId)
    {
        var self = XElement.Parse(oneNote.GetHierarchy(pageId, HierarchyScope.hsSelf));
        if (self.Name != One + "Page")
        {
            throw new McpException($"Not a page: {self.Name.LocalName}");
        }

        var section = SectionId().Match(pageId) is { Success: true } m ? TryPages($"{m.Groups[1].Value}{{1}}{{B0}}") : null;
        section = section is not null && Contains(section, pageId) ? section
            : TryPages("")?.Descendants(One + "Section").FirstOrDefault(s => Contains(s, pageId))
              ?? throw new McpException("The page's section wasn't found. Is its notebook open in OneNote desktop?");

        if (IsDeleted(section) || (string?)self.Attribute("isInRecycleBin") == "true")
        {
            throw new McpException("The page is in Deleted Pages, which OneNote keeps read-only");
        }

        return section;
    }

    /// <summary>The target section, with its pages.</summary>
    private XElement TargetSection(string sectionId)
    {
        var self = XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsSelf));
        if (self.Name != One + "Section")
        {
            throw new McpException($"sectionId must be a section, not a {self.Name.LocalName}");
        }

        if (IsDeleted(self))
        {
            throw new McpException("Can't move a page into Deleted Pages or the Recycle Bin. Use onenote_delete_page.");
        }

        return (string?)self.Attribute("isLocked") == "true"
            ? throw new McpException("The section is locked. Unlock it in OneNote desktop first.")
            : XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages));
    }

    /// <summary>Each moving page by its link page-id, which survives the move.</summary>
    /// <remarks>
    /// - Checked before the move: without a page-id, the page's new ID couldn't be found after it.
    /// </remarks>
    private Dictionary<string, XElement> MovingLinks(IReadOnlyList<XElement> moving)
    {
        var links = new Dictionary<string, XElement>();
        foreach (var page in moving)
        {
            if (LinkId(PageMoves.Id(page)) is not { } link || !links.TryAdd(link, page))
            {
                throw new McpException($"Can't move: OneNote gave \"{Name(page)}\" no unique link page-id, so its new ID couldn't be found after the move.");
            }
        }

        return links;
    }

    /// <summary>The moved pages' new IDs, matched by link page-id, in moving order.</summary>
    /// <remarks>
    /// - Stops once every page is matched, and skips pages whose link OneNote can't give, such as ones sync just added.
    /// </remarks>
    private List<MovedPage> NewIds(string sectionId, HashSet<string> before, Dictionary<string, XElement> links)
    {
        var found = new Dictionary<string, string>();
        foreach (var id in XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages)).Elements(One + "Page").Select(PageMoves.Id).Where(id => !before.Contains(id)))
        {
            if (found.Count == links.Count)
            {
                break;
            }

            if (TryLinkId(id) is { } link && links.ContainsKey(link))
            {
                found.TryAdd(link, id);
            }
        }

        return [.. links.Select(link => found.TryGetValue(link.Key, out var newId)
            ? new MovedPage(PageMoves.Id(link.Value), newId, Name(link.Value))
            : throw new McpException($"OneNote moved the page, but didn't report a new ID for \"{Name(link.Value)}\". Use onenote_list_pages on the target section."))];
    }

    private string? TryLinkId(string pageId)
    {
        try
        {
            return LinkId(pageId);
        }
        catch (McpException)
        {
            return null;
        }
    }

    private XElement? TryPages(string sectionId)
    {
        try
        {
            return XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages));
        }
        catch (McpException)
        {
            return null;
        }
    }

    /// <summary>A page's link page-id, which stays the same when the page moves. Null when the link has none.</summary>
    private string? LinkId(string pageId) =>
        LinkPageId().Match(oneNote.GetHyperlink(pageId)) is { Success: true } m ? m.Groups[1].Value.ToUpperInvariant() : null;

    private static bool Contains(XElement section, string pageId) =>
        section.Elements(One + "Page").Any(p => PageMoves.SameId(PageMoves.Id(p), pageId));

    private static string Name(XElement e) => (string?)e.Attribute("name") ?? "";

    [GeneratedRegex(@"^(\{[^}]+\})\{1\}\{E", RegexOptions.IgnoreCase)]
    private static partial Regex SectionId();

    /// <summary>The page-id in a link, with plain or URL-encoded braces</summary>
    [GeneratedRegex(@"page-id=(?:\{|%7B)([0-9A-Fa-f-]{36})(?:\}|%7D)", RegexOptions.IgnoreCase)]
    private static partial Regex LinkPageId();
}
