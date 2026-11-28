using System.Xml.Linq;
using ModelContextProtocol;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

/// <summary>
/// Places pages in a section's hierarchy XML, for <c>UpdateHierarchy</c> to move them.
/// </summary>
/// <remarks>
/// - Adding a page from another section, in any notebook, moves it there. The page gets a new ID, but its link's page-id stays.
/// - A section lists subpages as the pages after their parent with a deeper <c>pageLevel</c>.
/// </remarks>
public static class PageMoves
{
    /// <summary>A page and its subpages: the pages right after it with a deeper level.</summary>
    public static IReadOnlyList<XElement> WithSubpages(XElement page)
    {
        var level = Level(page);
        return [page, .. page.ElementsAfterSelf(One + "Page").TakeWhile(p => Level(p) > level)];
    }

    /// <summary>OneNote's deepest subpage level</summary>
    public const int MaxLevel = 3;

    /// <summary>A copy of the section's XML with the pages placed after <paramref name="afterPageId"/>, or at the end.</summary>
    /// <param name="pages">The page and its subpages, from <see cref="WithSubpages"/></param>
    /// <param name="afterPageId">A page in the section, or empty. The pages go after its subpages too, so those stay with it.</param>
    /// <remarks>
    /// - The moved page takes the level of the page it follows, so it becomes that page's sibling, or a top-level page at the end. Its subpages keep their depth under it.
    /// - Within one section, the pages leave their old place first.
    /// </remarks>
    public static XElement Place(XElement section, IReadOnlyList<XElement> pages, string afterPageId)
    {
        var moving = pages.Select(Id).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (moving.Contains(afterPageId))
        {
            throw new McpException("afterPageId can't be the page or one of its subpages");
        }

        var placed = new XElement(section);
        placed.Elements(One + "Page").Where(p => moving.Contains(Id(p))).Remove();
        var after = afterPageId.Length == 0 ? null
            : placed.Elements(One + "Page").FirstOrDefault(p => SameId(Id(p), afterPageId))
              ?? throw new McpException("afterPageId isn't a page in that section");

        var shift = (after is null ? 1 : Level(after)) - Level(pages[0]);
        if (pages.Max(Level) + shift > MaxLevel)
        {
            throw new McpException($"Its subpages would go deeper than OneNote's {MaxLevel} levels. Place it after a top-level page.");
        }

        var copies = pages.Select(p =>
        {
            var copy = new XElement(p);
            copy.SetAttributeValue("pageLevel", Level(p) + shift);
            return copy;
        }).ToList();

        if (after is null)
        {
            placed.Add(copies);
        }
        else
        {
            WithSubpages(after)[^1].AddAfterSelf(copies);
        }

        return placed;
    }

    public static string Id(XElement e) => (string?)e.Attribute("ID") ?? "";

    /// <summary>Whether two OneNote IDs match, ignoring case, as OneNote does.</summary>
    public static bool SameId(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private static int Level(XElement page) => int.TryParse((string?)page.Attribute("pageLevel"), out var level) ? level : 1;
}
