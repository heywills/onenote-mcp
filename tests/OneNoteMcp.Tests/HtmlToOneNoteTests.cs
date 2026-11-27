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
    [InlineData("<html><head><noscript><img src=\"https://x.test/pixel.gif\"></noscript></head><body><p>Body</p></body></html>")]
    [InlineData("<p>Body</p><noframes><p>Old browsers</p></noframes><iframe src=\"https://x.test\">Fallback</iframe>")]
    public void KeepsOnlyTheBodyContent(string html) =>
        Assert.Equal(["Body"], Texts(Convert(html)));

    [Fact]
    public void SkipsNonContentWithoutSplittingTheParagraph() =>
        Assert.Equal(["Hello world"], Texts(Convert("Hello <script>x()</script>world")));

    [Fact]
    public void SkipsNonContentInsideInlineText()
    {
        var page = Convert("<h1>Hi<script>track()</script></h1><p><a href=\"https://x.test\"><svg><title>Twitter</title></svg>Share</a></p>");
        Assert.Equal(["Hi", "<a href=\"https://x.test\">Share</a>"], Texts(page));
    }

    private static readonly byte[] Png = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private static (XElement Page, (int Embedded, int Linked) Images) ConvertWithImages(string html)
    {
        var page = new XElement(One + "Page", new XElement(One + "Title"));
        var images = HtmlToOneNote.AddToPage(page, html, loadImage: src => src.StartsWith("missing") ? null : Png);
        return (page, images);
    }

    // What each OE holds, in order: its text, or "image"
    private static List<string> Content(XElement page) => page.Descendants(One + "OE")
        .Select(oe => oe.Element(One + "Image") is not null ? "image" : (string?)oe.Element(One + "T") ?? "").ToList();

    [Fact]
    public void EmbedsImagesOnTheirOwnLine()
    {
        var (page, images) = ConvertWithImages("<p>before <img src=\"a.png\"> after</p><p><img src=\"missing.png\" alt=\"gone\"></p>");

        Assert.Equal(["before", "image", "after", "<a href=\"missing.png\">gone</a>"], Content(page));
        Assert.Equal((1, 1), images);
        var image = page.Descendants(One + "Image").Single();
        Assert.Equal("png", (string?)image.Attribute("format"));
        Assert.Equal(System.Convert.ToBase64String(Png), (string?)image.Element(One + "Data"));
    }

    [Fact]
    public void EmbedsImagesInListItems()
    {
        var (page, images) = ConvertWithImages("<ul><li><img src=\"a.png\"></li><li>text <img src=\"b.png\"></li></ul>");

        var items = page.Descendants(One + "OEChildren").First().Elements(One + "OE").ToList();
        Assert.NotNull(items[0].Element(One + "List"));
        Assert.NotNull(items[0].Element(One + "Image"));
        Assert.Null(items[0].Element(One + "T"));
        Assert.Equal("text", (string?)items[1].Element(One + "T"));
        Assert.NotNull(items[1].Element(One + "OEChildren")!.Element(One + "OE")!.Element(One + "Image"));
        Assert.Equal((2, 0), images);
    }

    [Fact]
    public void CountsOnlyShownImages() =>
        Assert.Equal((0, 1), ConvertWithImages("<noscript><img src=\"a.png\"></noscript><p><a href=\"x\"><img src=\"b.png\"></a></p>").Images);

    [Fact]
    public void LinksImagesPastThePageBudget()
    {
        var nineMb = new byte[9 * 1024 * 1024];
        Png.CopyTo(nineMb, 0);
        var page = new XElement(One + "Page", new XElement(One + "Title"));
        var html = string.Concat(Enumerable.Range(1, 6).Select(i => $"<p><img src=\"{i}.png\"></p>"));
        Assert.Equal((5, 1), HtmlToOneNote.AddToPage(page, html, loadImage: _ => nineMb));
    }

    [Fact]
    public void CountsLinksAsWritten() =>
        Assert.Equal((1, 2), ConvertWithImages("<h2>Logo <img src=\"missing-1.png\"></h2><ul><li><img src=\"a.png\"></li><li>text <img src=\"missing-2.png\"></li></ul>").Images);

    [Fact]
    public void KeepsImagesInAQuote()
    {
        var (page, _) = ConvertWithImages("<blockquote><p>Quoted</p><p><img src=\"a.png\"></p></blockquote>");
        var styles = page.Descendants(One + "OE").Select(oe => (string?)oe.Attribute("quickStyleIndex")).Distinct();
        Assert.Single(styles);
    }

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
