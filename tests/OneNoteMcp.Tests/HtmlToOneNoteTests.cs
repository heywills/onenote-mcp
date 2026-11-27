using System.Xml.Linq;

namespace OneNoteMcp.Tests;

public class HtmlToOneNoteTests
{
    private static readonly XNamespace One = Hierarchy.One;

    private static XElement Convert(string html)
    {
        var page = new XElement(One + "Page",
            new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "PageTitle")),
            new XElement(One + "Title"));
        HtmlToOneNote.AddToPage(page, html);
        return page;
    }

    private static List<string> Texts(XElement e) => e.Descendants(One + "T").Select(t => t.Value).ToList();

    [Theory]
    [InlineData("<!DOCTYPE html><html><head><title>Notes</title><style>p{color:red}</style></head><body><p>Body</p></body></html>")]
    [InlineData("<html><head><meta charset=\"utf-8\"><link rel=\"stylesheet\" href=\"s.css\"></head><body><p>Body</p><script>x()</script></body></html>")]
    [InlineData("<style>p{color:red}</style><p>Body</p><noscript>Enable JS</noscript>")]
    public void KeepsOnlyTheBodyContent(string html) =>
        Assert.Equal(["Body"], Texts(Convert(html)));

    [Fact]
    public void KeepsInlineFormattingAndSpaces()
    {
        var page = Convert("<p>Text with <b>bold</b>, <code>code</code>, and a <a href=\"https://x.test/?a=1&b=2\">link</a>.</p>");

        Assert.Equal(
            "Text with <b>bold</b>, <span style='font-family:Consolas'>code</span>, and a <a href=\"https://x.test/?a=1&amp;b=2\">link</a>.",
            Assert.Single(Texts(page)));
    }

    [Fact]
    public void AddsHeadingStylesAfterExistingStyles()
    {
        var page = Convert("<h1>One</h1><p>Body</p><h3>Three</h3>");

        var defs = page.Elements(One + "QuickStyleDef").ToList();
        Assert.Equal(["PageTitle", "h1", "h3"], defs.Select(d => (string?)d.Attribute("name")));
        Assert.Equal(One + "Title", defs[^1].ElementsAfterSelf().First().Name);

        var oes = page.Element(One + "Outline")!.Element(One + "OEChildren")!.Elements(One + "OE").ToList();
        Assert.Equal(["One", "Body", "", "Three"], oes.Select(o => o.Element(One + "T")!.Value));
        string StyleName(XElement oe) =>
            (string)defs.Single(d => (string?)d.Attribute("index") == (string?)oe.Attribute("quickStyleIndex")).Attribute("name")!;
        Assert.Equal("h1", StyleName(oes[0]));
        Assert.Equal("h3", StyleName(oes[3]));
    }

    [Fact]
    public void ReusesPageStylesAndTakesFreeIndexes()
    {
        var page = new XElement(One + "Page",
            new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "PageTitle")),
            new XElement(One + "QuickStyleDef", new XAttribute("index", "1"), new XAttribute("name", "p")),
            new XElement(One + "QuickStyleDef", new XAttribute("index", "2"), new XAttribute("name", "h2")));
        HtmlToOneNote.AddToPage(page, "<h1>One</h1><h2>Two</h2>");

        Assert.Equal(["0 PageTitle", "1 p", "2 h2", "3 h1"],
            page.Elements(One + "QuickStyleDef").Select(d => $"{d.Attribute("index")!.Value} {d.Attribute("name")!.Value}"));
        Assert.Equal(["3", "", "2"], page.Descendants(One + "OE").Select(o => (string?)o.Attribute("quickStyleIndex") ?? ""));
    }

    [Fact]
    public void KeepsImagesAsLinks() =>
        Assert.Equal("<a href=\"https://x.test/a.png\">Diagram</a>", Assert.Single(Texts(Convert("<p><img src=\"https://x.test/a.png\" alt=\"Diagram\"></p>"))));

    [Fact]
    public void NestsLists()
    {
        var page = Convert("<ul><li>Bullet<ul><li>Nested</li></ul></li></ul><ol><li>Step</li></ol>");

        var oes = page.Element(One + "Outline")!.Element(One + "OEChildren")!.Elements(One + "OE").ToList();
        Assert.NotNull(oes[0].Element(One + "List")!.Element(One + "Bullet"));
        var nested = oes[0].Element(One + "OEChildren")!.Element(One + "OE")!;
        Assert.Equal("Nested", nested.Element(One + "T")!.Value);
        Assert.NotNull(oes[1].Element(One + "List")!.Element(One + "Number"));
    }

    [Fact]
    public void BuildsTablesWithBoldHeaders()
    {
        var page = Convert("<table><tr><th>Name</th><th>Value</th></tr><tr><td>Beta</td><td>2 &amp; 3</td></tr><tr><td>Gamma</td></tr></table>");

        var table = page.Descendants(One + "Table").Single();
        Assert.Equal(2, table.Element(One + "Columns")!.Elements().Count());
        Assert.Equal(["<b>Name</b>", "<b>Value</b>", "Beta", "2 &amp; 3", "Gamma", ""], Texts(table));
    }

    [Fact]
    public void KeepsPreLinesAndSpaces()
    {
        var page = Convert("<pre>if (x)\n  y();</pre><pre>z();</pre>");

        Assert.Equal(["if&nbsp;(x)", "&nbsp;&nbsp;y();", "", "z();"], Texts(page));
        var code = page.Elements(One + "QuickStyleDef").Single(d => (string?)d.Attribute("name") == "code");
        Assert.Equal("Consolas", (string?)code.Attribute("font"));
        Assert.Equal(
            [code.Attribute("index")!.Value, code.Attribute("index")!.Value, null, code.Attribute("index")!.Value],
            page.Descendants(One + "OE").Select(o => (string?)o.Attribute("quickStyleIndex")));
    }
}
