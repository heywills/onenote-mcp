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

    /// <summary>A copy of the section's XML with the pages placed after <paramref name="afterPageId"/>, or at the end.</summary>
    /// <param name="pages">The page and its subpages, from <see cref="WithSubpages"/></param>
    /// <param name="afterPageId">A page in the section, or empty. The pages go after its subpages too, so those stay with it.</param>
    /// <remarks>
    /// - The moved page becomes a top-level page, and its subpages keep their depth under it.
    /// - Within one section, the pages leave their old place first.
    /// </remarks>
    public static XElement Place(XElement section, IReadOnlyList<XElement> pages, string afterPageId)
    {
        var moving = pages.Select(Id).ToHashSet();
        if (moving.Contains(afterPageId))
        {
            throw new McpException("afterPageId can't be the page or one of its subpages");
        }

        var placed = new XElement(section);
        placed.Elements(One + "Page").Where(p => moving.Contains(Id(p))).Remove();

        var shift = Level(pages[0]) - 1;
        var copies = pages.Select(p =>
        {
            var copy = new XElement(p);
            copy.SetAttributeValue("pageLevel", Level(p) - shift);
            return copy;
        }).ToList();

        if (string.IsNullOrEmpty(afterPageId))
        {
            placed.Add(copies);
        }
        else
        {
            var after = placed.Elements(One + "Page").FirstOrDefault(p => Id(p) == afterPageId)
                ?? throw new McpException("afterPageId isn't a page in that section");
            WithSubpages(after)[^1].AddAfterSelf(copies);
        }

        return placed;
    }

    public static string Id(XElement e) => (string?)e.Attribute("ID") ?? "";

    private static int Level(XElement page) => int.TryParse((string?)page.Attribute("pageLevel"), out var level) ? level : 1;
}
