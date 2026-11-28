using System.Xml.Linq;

namespace OneNoteMcp.Tests;

public class HtmlToOneNoteTests
{
    private static readonly XNamespace One = Hierarchy.One;

    /// <param name="fromMarkdown">Markdown's conventions: a blank line before headings, bold header cells, and hard breaks that split paragraphs</param>
    private static XElement Convert(string html, bool fromMarkdown = false)
    {
        var page = new XElement(One + "Page",
            new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "PageTitle")),
            new XElement(One + "Title"));
        HtmlToOneNote.AddToPage(page, html, fromMarkdown);
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
        var images = HtmlToOneNote.AddToPage(page, html, loadImage: (src, _) => src.StartsWith("missing") ? null : Png);
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
        var limits = new List<long>();
        Assert.Equal((5, 1), HtmlToOneNote.AddToPage(page, html, loadImage: (_, limit) =>
        {
            limits.Add(limit);
            return limit >= nineMb.Length ? nineMb : null;
        }));

        // With 5 MB left, the last read is capped there: a 9 MB file isn't read at all
        Assert.Equal(5 * 1024 * 1024, limits[^1]);
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
        var page = Convert("<h1>One</h1><p>Body</p><h3>Three</h3>", fromMarkdown: true);

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
        HtmlToOneNote.AddToPage(page, "<h1>One</h1><h2>Two</h2>", fromMarkdown: true);

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
        var page = Convert("<table><tr><th>Name</th><th>Value</th></tr><tr><td>Beta</td><td>2 &amp; 3</td></tr><tr><td>Gamma</td></tr></table>", fromMarkdown: true);

        var table = page.Descendants(One + "Table").Single();
        Assert.Equal(2, table.Element(One + "Columns")!.Elements().Count());
        Assert.Equal(["<b>Name</b>", "<b>Value</b>", "Beta", "2 &amp; 3", "Gamma", ""], Texts(table));
    }

    [Fact]
    public void KeepsHtmlFormatting()
    {
        var page = Convert("""
            <p style="font-family:Calibri;font-size:12.0pt;color:#14324D;text-align:center">Centered <span style='color:#FF0000;background:yellow;font-family:"Lucida Console";margin:0'>red</span> text<br>next line&nbsp;here</p>
            <h2 style="text-align:right">Right</h2>
            <table><tr><th>Head</th></tr><tr><td style="background-color:#DDEBF7">Shaded</td></tr></table>
            <p>Before</p><p></p><h3>After a blank</h3>
            """);

        var oe = page.Descendants(One + "OE").First();
        Assert.Equal("font-family:Calibri;font-size:12.0pt;color:#14324D", (string?)oe.Attribute("style"));
        Assert.Equal("center", (string?)oe.Attribute("alignment"));
        // Margin isn't text formatting, so it's dropped. The line break and non-breaking space stay.
        Assert.Equal("Centered <span style=\"color:#FF0000;background:yellow;font-family:&quot;Lucida Console&quot;\">red</span> text<br>\nnext line&#160;here",
            (string?)oe.Element(One + "T"));
        Assert.Equal("right", (string?)page.Descendants(One + "OE").Single(o => (string?)o.Element(One + "T") == "Right").Attribute("alignment"));

        var table = page.Descendants(One + "Table").Single();
        Assert.Equal("true", (string?)table.Attribute("hasHeaderRow"));
        Assert.Equal("Head", Texts(table)[0]);
        Assert.Equal("#DDEBF7", (string?)table.Descendants(One + "Cell").Last().Attribute("shadingColor"));

        // HTML keeps its own blank lines: one before the heading, not two
        Assert.Equal(["Before", "", "After a blank"], Texts(page).TakeLast(3));
    }

    [Fact]
    public void DropsScriptLinks() =>
        Assert.Equal(["x pic d m"],Texts(Convert("<p><a href=\"java&#9;script:alert(1)\">x</a> <img src=\"JavaScript:y\" alt=\"pic\"> <a href=\"data:text/html,x\">d</a> <a href=\"MS-MSDT:/id x\">m</a></p>")));

    [Fact]
    public void AddsNothingForAnEmptyIndent()
    {
        var page = Convert("<pre>A</pre><div class=\"indent\"></div><pre>C</pre>");
        Assert.DoesNotContain(page.Descendants(One + "OEChildren"), c => c.Parent!.Name == One + "OE");
        Assert.Equal(["A", "", "C"], Texts(page));
    }

    [Fact]
    public void KeepsCodeSplitByAnIndentAsOneBlock() =>
        Assert.Equal(["A", "B", "i", "C"], Texts(Convert("<pre>A\nB</pre><div class=\"indent\"><p>i</p></div><pre>C</pre>")));

    [Fact]
    public void KeepsAppLinks() =>
        Assert.Equal(["<a href=\"tel:+15551234\">Call</a> <a href=\"msteams:/l/chat\">Chat</a>"],
            Texts(Convert("<p><a href=\"tel:+15551234\">Call</a> <a href=\"msteams:/l/chat\">Chat</a></p>")));

    [Fact]
    public void WritesImagesWithoutASourceAsText() =>
        Assert.Equal(["Image 1"], Texts(Convert("<p><img data-callback-id=\"cb\" alt=\"Image 1\"></p>")));

    [Fact]
    public void KeepsTheInnerParagraphsAlignment() =>
        Assert.Equal("center", (string?)Convert("<div style=\"text-align:left\"><p style=\"text-align:center\">C</p></div>").Descendants(One + "OE").Single().Attribute("alignment"));

    [Fact]
    public void MakesRowHeadersBoldText()
    {
        var page = Convert("<table><tr><th>Name</th><td>Ann</td></tr></table>");
        Assert.Null(page.Descendants(One + "Table").Single().Attribute("hasHeaderRow"));
        Assert.Equal(["<b>Name</b>", "Ann"], Texts(page));
    }

    [Fact]
    public void KeepsOnlyColorFunctions() =>
        Assert.Equal(["<span style=\"color:rgb(1,2,3)\">s</span>"],
            Texts(Convert("<p><span style=\"color:rgb(1,2,3);background:image-set(&quot;https://x.test/a.png&quot; 1x);font-family:url/**/(x)\">s</span></p>")));

    [Fact]
    public void SkipsLineBreaksBetweenBlocks() =>
        Assert.Equal(["Intro", "Next", ""], Texts(Convert("<div>Intro</div><br><h2>Next</h2>\n<br>\n<p><br></p>")));

    [Fact]
    public void BoldsHeaderCellsOutsideTheHeaderRow() =>
        Assert.Equal(["H", "", "<b>Row</b>", "v"], Texts(Convert("<table><tr><th>H</th><th></th></tr><tr><th>Row</th><td>v</td></tr></table>")));

    [Fact]
    public void MakesPositionedDivsIntoOutlines()
    {
        var page = Convert("""
            <p>Loose</p>
            <div class="outline" style="position:absolute;left:0pt;top:86.4pt;width:540pt;z-index:5"><p>Main</p></div>
            <div class="outline" style="position:absolute;left:612pt;top:-3pt"><p>Note</p></div>
            <div class="outline"><p>Unplaced stays in order</p></div>
            <div class="outline" style="position:absolute;left:NaNpt;top:1e9pt"><p>Bad place</p></div>
            <div class="outline" style="position:absolute;left:-0.001pt;top:10pt;width:0.004pt;z-index:2147483647"><p>Edge</p></div>
            """);

        var outlines = page.Elements(One + "Outline").ToList();
        Assert.Equal(["LooseUnplaced stays in orderBad place", "Main", "Note", "Edge"], outlines.Select(o => string.Concat(Texts(o))));
        // A tiny negative place is 0, a width that rounds to 0 is dropped, and a huge z-index falls back to the order
        Assert.Equal(("0", "10", "8"), Position(outlines[3]));
        Assert.Null(outlines[3].Element(One + "Size"));
        Assert.Null(outlines[0].Element(One + "Position"));
        Assert.Equal(("0", "86.4", "5"), Position(outlines[1]));
        Assert.Equal("540", outlines[1].Element(One + "Size")!.Attribute("width")!.Value);
        // Without a z-index, above the highest one given
        Assert.Equal(("612", "-3", "7"), Position(outlines[2]));
        Assert.Null(outlines[2].Element(One + "Size"));
    }

    private static (string, string, string) Position(XElement outline) =>
        (outline.Element(One + "Position")!.Attribute("x")!.Value, outline.Element(One + "Position")!.Attribute("y")!.Value, outline.Element(One + "Position")!.Attribute("z")!.Value);

    [Fact]
    public void SizesImagesFromPoints()
    {
        var page = new XElement(One + "Page", new XElement(One + "Title"));
        HtmlToOneNote.AddToPage(page, "<p><img src=\"a.png\" style=\"width:120.5pt;height:60pt\"></p>", loadImage: (_, _) => Png);
        var size = page.Descendants(One + "Size").Single();
        Assert.Equal(("120.5", "60"), ((string?)size.Attribute("width"), (string?)size.Attribute("height")));
        Assert.Equal(One + "Data", size.ElementsAfterSelf().Single().Name);
    }

    [Fact]
    public void NestsIndentedBlocksUnderTheParagraphBefore()
    {
        var page = Convert("<p>Parent</p><div class=\"indent\"><p>Child</p></div>");
        var parent = page.Descendants(One + "OE").First();
        Assert.Equal("Child", (string?)parent.Element(One + "OEChildren")!.Element(One + "OE")!.Element(One + "T"));
    }

    [Fact]
    public void ReadsTheTitle()
    {
        Assert.Equal("My page", HtmlToOneNote.Title("<html><head><title> My page </title></head><body></body></html>"));
        Assert.Null(HtmlToOneNote.Title("<p>No title</p>"));
    }

    [Fact]
    public void AddsNoOutlineForEmptyContent()
    {
        var page = Convert("<!DOCTYPE html><html><head><title>Empty</title></head><body></body></html>");
        Assert.Empty(page.Elements(One + "Outline"));
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
