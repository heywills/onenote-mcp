using System.Xml.Linq;
using ModelContextProtocol;

namespace OneNoteMcp.Tests;

/// <summary>
/// Placing a page and its subpages in a section's hierarchy XML.
/// </summary>
public class PageMoveTests
{
    private static readonly XNamespace One = Hierarchy.One;

    /// <summary>A section whose pages are (id, level) pairs.</summary>
    private static XElement Section(params (string Id, int Level)[] pages) =>
        new(One + "Section", new XAttribute("ID", "s"),
            pages.Select(p => new XElement(One + "Page", new XAttribute("ID", p.Id), new XAttribute("name", p.Id), new XAttribute("pageLevel", p.Level))));

    private static List<string> Order(XElement section) =>
        [.. section.Elements(One + "Page").Select(p => $"{p.Attribute("ID")!.Value}:{p.Attribute("pageLevel")!.Value}")];

    private static XElement Page(XElement section, string id) => section.Elements(One + "Page").Single(p => p.Attribute("ID")!.Value == id);

    [Fact]
    public void TakesThePageWithItsSubpages()
    {
        var source = Section(("a", 1), ("a1", 2), ("a1x", 3), ("b", 1), ("b1", 2));
        Assert.Equal(["a", "a1", "a1x"], PageMoves.WithSubpages(Page(source, "a")).Select(PageMoves.Id));
        Assert.Equal(["a1", "a1x"], PageMoves.WithSubpages(Page(source, "a1")).Select(PageMoves.Id));
    }

    [Fact]
    public void PlacesAtTheEndAsATopLevelPage()
    {
        var source = Section(("p", 1), ("a", 2), ("a1", 3));
        var placed = PageMoves.Place(Section(("x", 1)), PageMoves.WithSubpages(Page(source, "a")), "");

        Assert.Equal(["x:1", "a:1", "a1:2"], Order(placed));
    }

    [Fact]
    public void PlacesAfterAPageAndItsSubpages()
    {
        var source = Section(("a", 1), ("a1", 2));
        var placed = PageMoves.Place(Section(("x", 1), ("x1", 2), ("y", 1)), PageMoves.WithSubpages(Page(source, "a")), "x");

        Assert.Equal(["x:1", "x1:2", "a:1", "a1:2", "y:1"], Order(placed));
    }

    [Fact]
    public void ReordersWithinASection()
    {
        var section = Section(("a", 1), ("a1", 2), ("b", 1), ("c", 1));
        var placed = PageMoves.Place(section, PageMoves.WithSubpages(Page(section, "a")), "c");

        Assert.Equal(["b:1", "c:1", "a:1", "a1:2"], Order(placed));
        // The input stays as it was
        Assert.Equal(["a:1", "a1:2", "b:1", "c:1"], Order(section));
    }

    [Theory]
    [InlineData("a1", "can't be the page or one of its subpages")]
    [InlineData("nope", "isn't a page in that section")]
    public void RefusesABadPlace(string afterPageId, string error)
    {
        var section = Section(("a", 1), ("a1", 2), ("b", 1));
        Assert.Contains(error, Assert.Throws<McpException>(() => PageMoves.Place(section, PageMoves.WithSubpages(Page(section, "a")), afterPageId)).Message);
    }
}
