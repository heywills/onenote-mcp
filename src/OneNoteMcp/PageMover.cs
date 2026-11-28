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

        // Each page's link page-id, which survives the move
        var links = moving.ToDictionary(p => LinkId(PageMoves.Id(p)), p => p);
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
            : XElement.Parse(oneNote.GetHierarchy("", HierarchyScope.hsPages)).Descendants(One + "Section").FirstOrDefault(s => Contains(s, pageId))
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

        return IsDeleted(self)
            ? throw new McpException("Can't move a page into Deleted Pages or the Recycle Bin. Use onenote_delete_page.")
            : XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages));
    }

    /// <summary>The moved pages' new IDs, matched by link page-id, in moving order.</summary>
    private List<MovedPage> NewIds(string sectionId, HashSet<string> before, Dictionary<string, XElement> links)
    {
        var found = XElement.Parse(oneNote.GetHierarchy(sectionId, HierarchyScope.hsPages)).Elements(One + "Page")
            .Select(PageMoves.Id)
            .Where(id => !before.Contains(id))
            .ToDictionary(LinkId, id => id);

        return [.. links.Select(link => found.TryGetValue(link.Key, out var newId)
            ? new MovedPage(PageMoves.Id(link.Value), newId, Name(link.Value))
            : throw new McpException($"OneNote moved the page, but didn't report a new ID for \"{Name(link.Value)}\". Use onenote_list_pages on the target section."))];
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

    /// <summary>A page's link page-id, which stays the same when the page moves.</summary>
    private string LinkId(string pageId) =>
        LinkPageId().Match(oneNote.GetHyperlink(pageId)) is { Success: true } m ? m.Groups[1].Value.ToUpperInvariant() : pageId;

    private static bool Contains(XElement section, string pageId) =>
        section.Elements(One + "Page").Any(p => PageMoves.SameId(PageMoves.Id(p), pageId));

    private static string Name(XElement e) => (string?)e.Attribute("name") ?? "";

    [GeneratedRegex(@"^(\{[^}]+\})\{1\}\{E", RegexOptions.IgnoreCase)]
    private static partial Regex SectionId();

    [GeneratedRegex(@"page-id=\{([0-9A-Fa-f-]{36})\}")]
    private static partial Regex LinkPageId();
}
