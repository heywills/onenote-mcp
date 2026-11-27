using System.Xml.Linq;

namespace OneNoteMcp.Tests;

public class MarkdownTests
{
    private const string Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    internal const string Sample = """
        Intro with **bold**, *italic*, ~~struck~~, `inline code`, and a [link](https://example.com).

        ## Lists

        - Bullet one
        - Bullet two
          - Nested bullet
            - Deeper bullet
        - Bullet three

        1. First
        2. Second
           1. Second A
           2. Second B
        3. Third

        ## Table

        | Name | Value |
        | --- | --- |
        | Alpha | 1 |
        | Beta | 2 & 3 |

        ### Code

        ```
        function hello() {

            return "world";
        }
        ```

        Last paragraph.
        """;

    private static string RoundTrip(string markdown)
    {
        var page = new XElement(XName.Get("Page", Ns), new XAttribute("name", "Sample"),
            new XElement(XName.Get("QuickStyleDef", Ns), new XAttribute("index", "0"), new XAttribute("name", "PageTitle")),
            new XElement(XName.Get("Title", Ns)));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml(markdown), fromMarkdown: true);
        return PageMarkdown.Convert(page.ToString()).Markdown;
    }

    [Fact]
    public void RoundTripsMarkdownThroughOneNoteXml() =>
        Assert.Equal($"# Sample\n\n{Sample}", RoundTrip(Sample));

    [Theory]
    [InlineData(@"Literal \*stars\* and \`ticks\` and a \\ backslash")]
    [InlineData(@"\_edge\_ underscores, but snake_case stays")]
    [InlineData(@"\# not a heading")]
    [InlineData(@"1\. not a list")]
    [InlineData(@"\- not a bullet")]
    [InlineData(@"\> not a quote")]
    [InlineData("2 < 3 & 4 > 1")]
    [InlineData("| A | B |\n| --- | --- |\n| x \\| y | z |")]
    [InlineData("***bold italic*** and **[bold link](https://x.test/?a=1&b=2)**")]
    [InlineData("> Quoted line\n\nAfter the quote")]
    [InlineData("- [ ] Task item")]
    [InlineData("`npm install`")]
    [InlineData("## `Config`")]
    [InlineData("- `foo`")]
    [InlineData("Run ``a `tick` b`` now")]
    [InlineData("```\na\n```\n\n```\nb\n```")]
    [InlineData("> *a* and *b*")]
    [InlineData("> - item one\n> - item two\n\nAfter")]
    [InlineData("> First paragraph\n>\n> Second paragraph")]
    [InlineData("> - quoted item\n\n> Separate quote")]
    [InlineData(@"\>note")]
    [InlineData(@"\---")]
    [InlineData(@"- 1\. step")]
    [InlineData(@"\~\~not struck\~\~ and see [1\](x) and \<b>tag")]
    [InlineData("x^2^ and ~sub~ and ==mark== and ++ins++")]
    [InlineData("a\n\n---\n\nb")]
    [InlineData("- item\n\n  > quote in the item\n\n- next")]
    [InlineData("````\n```\nx\n```\n````\n\nafter")]
    [InlineData(@"[1\]: https://example.com")]
    [InlineData(@"## Ticket \#")]
    [InlineData("#### Four\n\n###### Six")]
    [InlineData("- a\n\n  > - b\n  >   - c")]
    [InlineData(@"Use \&nbsp; or \&copy; but keep R&D")]
    public void RoundTripsEdgeCases(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    // OneNote: numberSequence picks the style (1 is I, II), and numbering continues while the format stays the same.
    [Fact]
    public void AlternatesNumberFormatsSoEachListRestarts()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("1. A\n2. B\n\nText\n\n1. C\n   1. Nested\n\nText\n\n1. D"));

        var numbers = page.Descendants(XName.Get("Number", Ns)).ToList();
        Assert.All(numbers, n => Assert.Equal("0", (string?)n.Attribute("numberSequence")));
        Assert.Equal(["##.", "##.", "##)", "##.", "##."], numbers.Select(n => (string?)n.Attribute("numberFormat")));
    }

    [Fact]
    public void KeepsWholeItalicInAStyleThatIsNotItalic() =>
        Assert.Equal("# Sample\n\n> *whole quote italic*", RoundTrip("> *whole quote italic*"));

    [Theory]
    [InlineData("1. a\n\n1) b")]
    [InlineData("- a\n\n* b")]
    [InlineData("1. a\n\nText\n\n1. b\n\nText\n\n1. c")]
    [InlineData("`x  =  1`")]
    public void RoundTripsListsAndCodeSpaces(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Fact]
    public void MergesNeighborRunsWithTheSameFormatting() =>
        Assert.Equal("# P\n\n**foobar** and *x y*", Read(
            "<one:OE><one:T style=\"font-weight:bold\"><![CDATA[foo]]></one:T><one:T style=\"font-weight:bold\"><![CDATA[bar]]></one:T>"
            + "<one:T><![CDATA[ and <span style='font-style:italic'>x </span><span style='font-style:italic'>y</span>]]></one:T></one:OE>"));

    [Fact]
    public void ListItemWithoutTextKeepsItsMarkerAndNesting()
    {
        const string bullet = "<one:List><one:Bullet bullet=\"2\"/></one:List>";
        var markdown = Read(
            $"<one:OE>{bullet}<one:T>a</one:T></one:OE><one:OE>{bullet}<one:Image><one:CallbackID callbackID=\"cb\"/></one:Image></one:OE><one:OE>{bullet}<one:T>c</one:T></one:OE>");

        // No blank line after a bare marker: "-\n\n  x" would end the item before x.
        Assert.Equal("# P\n\n- a\n-\n  [Image 1]\n\n- c", markdown);
        Assert.Equal("<ul>\n<li><p>a</p>\n</li>\n<li><p>[Image 1]</p>\n</li>\n<li><p>c</p>\n</li>\n</ul>",
            MarkdownToOneNote.ToHtml(markdown.Split("\n\n", 2)[1]).Trim());

        Assert.Equal("# P\n\n1.\n   child", Read(
            "<one:OE><one:List><one:Number numberSequence=\"0\" numberFormat=\"##.\" text=\"1.\"/></one:List><one:OEChildren><one:OE><one:T><![CDATA[child]]></one:T></one:OE></one:OEChildren></one:OE>"));

        // A truly empty item keeps the blank line before what follows
        Assert.Equal("# P\n\n- a\n-\n\nafter", Read($"<one:OE>{bullet}<one:T>a</one:T></one:OE><one:OE>{bullet}<one:T></one:T></one:OE><one:OE><one:T>after</one:T></one:OE>"));
    }

    [Fact]
    public void RepeatsThePlaceholderForAnImageShownAgain() =>
        Assert.Equal("# P\n\n[Image 1]\n\nbetween\n\n[Image 1]", Read(
            "<one:OE><one:Image><one:CallbackID callbackID=\"cb\"/></one:Image></one:OE><one:OE><one:T>between</one:T></one:OE><one:OE><one:Image><one:CallbackID callbackID=\"cb\"/></one:Image></one:OE>"));

    [Fact]
    public void UsesHtmlWhereEmphasisMarkersCannotOpenOrClose() =>
        Assert.Equal("# P\n\nfunc<strong>(x)</strong> and <strong>\"Note\"</strong>s but **bold** (**ok**)", Read(
            "<one:OE><one:T><![CDATA[func<span style='font-weight:bold'>(x)</span> and <span style='font-weight:bold'>\"Note\"</span>s but <b>bold</b> (<b>ok</b>)]]></one:T></one:OE>"));

    [Theory]
    [InlineData("func<strong>(x)</strong> and <em>\"Note\"</em>s")]
    [InlineData("[`a[0]`](https://x) and [a\\\\\\]](https://y)")]
    public void RoundTripsEmphasisAndLinkEdgeCases(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Fact]
    public void EscapesBracketsInLinkTextOutsideCode() =>
        Assert.Equal("# P\n\n[`a[0]`](https://x) and [a\\\\\\]](https://y)", Read(
            "<one:OE><one:T><![CDATA[<a href=\"https://x\"><span style='font-family:Consolas'>a[0]</span></a> and <a href=\"https://y\">a\\]</a>]]></one:T></one:OE>"));

    [Fact]
    public void TrimsByteOrderMarkBeforeConverting() =>
        Assert.StartsWith("<h1>", MarkdownToOneNote.ToHtml((char)0xFEFF + "# Heading"));

    [Fact]
    public void TableInsideAQuoteDoesNotAddTheQuoteStyle()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("> | A |\n> | --- |\n> | 1 |"));

        Assert.Empty(page.Elements(XName.Get("QuickStyleDef", Ns)));
    }

    private static string Read(string body, string pageName = "P", bool hasTitle = true) =>
        PageMarkdown.Convert($"""<one:Page xmlns:one="{Ns}" name="{pageName}">{(hasTitle ? "<one:Title/>" : "")}<one:Outline><one:OEChildren>{body}</one:OEChildren></one:Outline></one:Page>""").Markdown;

    [Fact]
    public void JoinsTextRunsOfOneParagraph()
    {
        const string bullet = "<one:List><one:Bullet bullet=\"2\"/></one:List>";
        Assert.Equal("# P\n\n- Hello world", Read($"<one:OE>{bullet}<one:T><![CDATA[Hello ]]></one:T><one:T selected=\"all\"><![CDATA[]]></one:T><one:T><![CDATA[world]]></one:T></one:OE>"));
        Assert.Equal("# P\n\nPlain **bold**", Read("<one:OE><one:T><![CDATA[Plain ]]></one:T><one:T style=\"font-weight:bold\"><![CDATA[bold]]></one:T></one:OE>"));
    }

    [Fact]
    public void KeepsChildrenOfTheRepeatedTitleLine() =>
        Assert.Equal("# Meeting\n\n- child bullet", Read(
            "<one:OE><one:T><![CDATA[Meeting]]></one:T><one:OEChildren><one:OE><one:List><one:Bullet bullet=\"2\"/></one:List><one:T><![CDATA[child bullet]]></one:T></one:OE></one:OEChildren></one:OE>",
            "Meeting", hasTitle: false));

    // The way OneNote stores h4 and h6: the style's italic copied onto the text, as spans or OE style
    [Fact]
    public void DropsStyleItalicFromH4AndH6()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="P"><one:Title/>
              <one:QuickStyleDef index="4" name="h4" italic="true"/><one:QuickStyleDef index="6" name="h6" italic="true"/>
              <one:Outline><one:OEChildren>
                <one:OE quickStyleIndex="4"><one:T><![CDATA[<span style='font-style:italic'>Four</span>]]></one:T></one:OE>
                <one:OE quickStyleIndex="6" style="font-style:italic"><one:T><![CDATA[Six]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal("# P\n\n#### Four\n\n###### Six", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void KeepsSideBySideListsApart()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="P"><one:Title/>
              <one:Outline><one:OEChildren><one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[left]]></one:T></one:OE></one:OEChildren></one:Outline>
              <one:Outline><one:OEChildren><one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[right]]></one:T></one:OE></one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal("# P\n\n- left\n\n* right", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void WritesBareUrlsAsAutolinks() =>
        Assert.Equal("# P\n\nSee <https://x.test/a_b> and <https://x.test/_x_/a>", Read(
            "<one:OE><one:T><![CDATA[See <a href=\"https://x.test/a_b\">https://x.test/a_b</a> and <a href=\"https://x.test/_x_/a\">https://x.test/_x_/a</a>]]></one:T></one:OE>"));

    [Theory]
    [InlineData(@"[Spec](\\\\fileserver\\share\\spec.docx) and [C](file:///C:\\_x)")]
    [InlineData("<https://x.test/_x_/a>")]
    [InlineData("```\n```\n\nafter\n\n```\n\nx\n\n```")]
    public void RoundTripsLinksAndCodeBlankLines(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Fact]
    public void MergesRunsThatDifferOnlyInColor() =>
        Assert.Equal("# P\n\n**foobar** and *xy*", Read(
            "<one:OE><one:T style=\"font-weight:bold\"><![CDATA[foo]]></one:T><one:T style=\"font-weight:bold;color:red\"><![CDATA[bar]]></one:T>"
            + "<one:T><![CDATA[ and <i>x</i><span style='font-style:italic;font-family:Calibri'>y</span>]]></one:T></one:OE>"));

    [Fact]
    public void SeesTheCharacterBeforeAWrapperSpan() =>
        Assert.Equal("# P\n\nfunc<strong>(x)</strong>", Read(
            "<one:OE><one:T><![CDATA[func<span style='color:red'><span style='font-weight:bold'>(x)</span></span>]]></one:T></one:OE>"));

    [Theory]
    [InlineData("- item\n\n  ```\n  a\n  ```\n\n```\nb\n```")]
    [InlineData("- x\n  - a\n\n  * b")]
    [InlineData("> - a\n>\n> * b")]
    [InlineData("- x\n\n> - quoted")]
    public void RoundTripsNestedBlocks(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Theory]
    [InlineData("```\nfunc f() {\n\treturn 1\n}\n```")]
    [InlineData("Run `a\tb` now")]
    public void KeepsTabsInCode(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Fact]
    public void PutsABlankLineBeforeANestedListThatStartsAboveOne()
    {
        const string bullet = "<one:List><one:Bullet bullet=\"2\"/></one:List>";
        static string Num(string text) => $"<one:List><one:Number numberSequence=\"0\" numberFormat=\"##.\" text=\"{text}\"/></one:List>";
        var markdown = Read($"<one:OE>{bullet}<one:T>B</one:T><one:OEChildren><one:OE>{Num("3.")}<one:T>y</one:T></one:OE><one:OE>{Num("4.")}<one:T>z</one:T></one:OE></one:OEChildren></one:OE>"
            + $"<one:OE>{bullet}<one:T>C</one:T></one:OE>");

        Assert.Equal("# P\n\n- B\n\n  3. y\n  4. z\n- C", markdown);
        Assert.Contains("<ol start=\"3\">", MarkdownToOneNote.ToHtml(markdown));
    }

    [Fact]
    public void DoesNotWrapFormattingTheParagraphAlreadyHas() =>
        Assert.Equal("# P\n\n**x y**", Read("<one:OE style=\"font-weight:bold\"><one:T><![CDATA[<span style='font-weight:bold'>x</span> y]]></one:T></one:OE>"));

    [Fact]
    public void LooksPastEmptyElementsForTheNextCharacter() =>
        Assert.Equal("# P\n\n<strong>foo.</strong>bar", Read("<one:OE><one:T><![CDATA[<span style='font-weight:bold'>foo.</span><span></span>bar]]></one:T></one:OE>"));

    [Theory]
    [InlineData("| H |\n| --- |\n| `a\\|b` and c\\|d |")]
    [InlineData("**`x`** and `a`**`b`**")]
    public void RoundTripsPipesAndFormattingInCode(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(RoundTrip(markdown)["# Sample\n\n".Length..]));

    [Theory]
    [InlineData("- a\n\n  b\n\n* c")]
    [InlineData("> - item\n>\n>   quoted child")]
    public void RoundTripsListEndingsAndQuotedChildren(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Fact]
    public void KeepsAPipeInHtmlTableCodeAsWritten()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, "<table><tr><td><code>grep a\\|b</code></td></tr></table>");

        Assert.Equal("<span style='font-family:Consolas'>grep&nbsp;a\\|b</span>", page.Descendants(XName.Get("T", Ns)).Last().Value);
    }

    [Fact]
    public void ReadsQuotedItemChildrenAndCellCodeAndLinksInCode()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="P"><one:Title/>
              <one:QuickStyleDef index="5" name="blockquote"/><one:QuickStyleDef index="6" name="code"/>
              <one:Outline><one:OEChildren>
                <one:OE quickStyleIndex="5"><one:List><one:Bullet bullet="2"/></one:List><one:T>item</one:T>
                  <one:OEChildren><one:OE><one:Image><one:CallbackID callbackID="cb"/></one:Image></one:OE></one:OEChildren></one:OE>
              </one:OEChildren></one:Outline>
              <one:Outline><one:OEChildren>
                <one:OE><one:Table>
                  <one:Row><one:Cell><one:OEChildren><one:OE><one:T>H</one:T></one:OE></one:OEChildren></one:Cell></one:Row>
                  <one:Row><one:Cell><one:OEChildren><one:OE quickStyleIndex="6"><one:T>a*b</one:T></one:OE></one:OEChildren></one:Cell></one:Row>
                </one:Table></one:OE>
                <one:OE><one:T><![CDATA[<span style='font-family:Consolas'>run <a href="https://x">cmd</a></span> and <b>Note:</b><br>next]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal("# P\n\n> - item\n>\n>   [Image 1]\n\n| H |\n| --- |\n| `a*b` |\n\n`run` [`cmd`](https://x) and **Note:** next", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void UsesHtmlForEmphasisRightAfterAnotherRunsMarker() =>
        Assert.Equal("# P\n\n*foo*<strong>.bar</strong>", Read("<one:OE><one:T><![CDATA[<span style='font-style:italic'>foo</span><span style='font-weight:bold'>.bar</span>]]></one:T></one:OE>"));

    [Fact]
    public void KeepsLinksInMonospaceText()
    {
        Assert.Equal("# P\n\n`see` [`doc`](http://x.com/)", Read("<one:OE><one:T><![CDATA[<span style='font-family:Consolas'>see <a href=\"http://x.com/\">doc</a></span>]]></one:T></one:OE>"));
        Assert.Equal("# P\n\n[`a*b_`](http://x.com/)", Read("<one:OE><one:T><![CDATA[<a href=\"http://x.com/\" style='font-family:Consolas'>a*b_</a>]]></one:T></one:OE>"));
    }

    [Fact]
    public void InlineCodeWithANewlineStaysOneLine()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, "<p><code>a\nb</code></p>");

        Assert.Equal("<span style='font-family:Consolas'>a&nbsp;b</span>", page.Descendants(XName.Get("T", Ns)).Single().Value);
    }

    [Fact]
    public void WritesOneNotePageLinksInTheCopyLinkForm()
    {
        const string stored = "onenote:#Agent%20tool%20comparisons&section-id=%7BC6D0BF8C-A2F7-40C0-B0EC-AE82CCC4E004%7D&page-id=%7B070EF062-FF02-48B1-9B99-08B45AD94F1D%7D&end&base-path=https://x.sharepoint.com/OneNote/Mike%20@%20Work/Agent%20tools.one";
        var markdown = Read($"<one:OE><one:T><![CDATA[<a href=\"{stored}\">Page</a>]]></one:T></one:OE>");

        Assert.Equal("# P\n\n[Page](<onenote:https://x.sharepoint.com/OneNote/Mike%20@%20Work/Agent%20tools.one#Agent tool comparisons&section-id={C6D0BF8C-A2F7-40C0-B0EC-AE82CCC4E004}&page-id={070EF062-FF02-48B1-9B99-08B45AD94F1D}&end>)", markdown);
        Assert.NotNull(PageLink.Parse(markdown));
    }

    [Fact]
    public void ReadsBoldCodeAndMergesCodeRuns()
    {
        Assert.Equal("# P\n\n**`x`**", Read("<one:OE><one:T><![CDATA[<span style='font-family:Consolas;font-weight:bold'>x</span>]]></one:T></one:OE>"));
        Assert.Equal("# P\n\n`ab`", Read("<one:OE><one:T><![CDATA[<span style='font-family:Consolas'>a</span><span style='font-family:Consolas;font-weight:bold'>b</span>]]></one:T></one:OE>"));
    }

    [Fact]
    public void NestedListsUnderDifferentItemsKeepTheFirstFormat()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("1. a\n   1. x\n2. b\n   1. y"));

        Assert.Equal(["##.", "##.", "##.", "##."], page.Descendants(XName.Get("Number", Ns)).Select(n => (string?)n.Attribute("numberFormat")));
    }

    [Fact]
    public void FindsTheTitleAfterACommentAndKeepsTheComment()
    {
        var (title, body) = MarkdownToOneNote.SplitTitle("<!-- c -->\n# Title\n\nBody");

        Assert.Equal("Title", title);
        Assert.Equal("<!-- c -->\n\nBody", body);
    }

    [Fact]
    public void ReadsAPageWithADuplicateStyleIndex() =>
        Assert.Equal("# P\n\nx", PageMarkdown.Convert($"""<one:Page xmlns:one="{Ns}" name="P"><one:Title/><one:QuickStyleDef index="1" name="p"/><one:QuickStyleDef index="1" name="h1"/><one:Outline><one:OEChildren><one:OE quickStyleIndex="1"><one:T>x</one:T></one:OE></one:OEChildren></one:Outline></one:Page>""").Markdown);

    [Fact]
    public void RestoresBackslashesOnlyInFileLinks()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml(@"[f](https://x.test/files/a%5Cb.txt) [g](file:///C:\\x)"));

        Assert.Equal("<a href=\"https://x.test/files/a%5Cb.txt\">f</a> <a href=\"file:///C:\\x\">g</a>", page.Descendants(XName.Get("T", Ns)).Single().Value);
    }

    [Fact]
    public void SkipsStyleDefsWithANonNumericIndex()
    {
        var page = new XElement(XName.Get("Page", Ns), new XElement(XName.Get("QuickStyleDef", Ns), new XAttribute("index", "x"), new XAttribute("name", "p")));
        HtmlToOneNote.AddToPage(page, "<h1>H</h1>");

        Assert.Contains(page.Elements(XName.Get("QuickStyleDef", Ns)), d => (string?)d.Attribute("name") == "h1");
    }

    [Fact]
    public void TitleKeepsAnAutolink() =>
        Assert.Equal("Notes for https://x.test", MarkdownToOneNote.SplitTitle("# Notes for <https://x.test>").Title);

    [Fact]
    public void WritesAnchorsWithoutAHrefAsText() =>
        Assert.Equal("# P\n\nTop of page", Read("<one:OE><one:T><![CDATA[<a name=\"top\">Top</a> of page]]></one:T></one:OE>"));

    [Fact]
    public void LinkedImageBecomesItsAltTextInsideTheLink()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("[![logo](https://x/l.png)](https://x)"));

        Assert.Equal("<a href=\"https://x\">logo</a>", page.Descendants(XName.Get("T", Ns)).Single().Value);
    }

    [Theory]
    [InlineData("Issue #", @"Issue \#")]
    [InlineData("*Draft* notes", @"\*Draft\* notes")]
    [InlineData("Using <b> tags", @"Using \<b> tags")]
    public void EscapesThePageTitle(string title, string heading)
    {
        var markdown = PageMarkdown.Convert($"""<one:Page xmlns:one="{Ns}" name="{System.Security.SecurityElement.Escape(title)}"><one:Title/></one:Page>""").Markdown;

        Assert.Equal($"# {heading}", markdown);
        Assert.Equal(title, MarkdownToOneNote.SplitTitle(markdown).Title);
    }

    [Fact]
    public void JoinsTextRunsInTableCells()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="P"><one:Title/><one:Outline><one:OEChildren><one:OE><one:Table>
              <one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[H]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>
              <one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[plain ]]></one:T><one:T style="font-weight:bold"><![CDATA[bold]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>
            </one:Table></one:OE></one:OEChildren></one:Outline></one:Page>
            """;

        Assert.Equal("# P\n\n| H |\n| --- |\n| plain **bold** |", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void ResetsAlternateMarkersAfterANonListOutline()
    {
        static string Outline(string body) => $"<one:Outline><one:OEChildren>{body}</one:OEChildren></one:Outline>";
        const string bullet = "<one:List><one:Bullet bullet=\"2\"/></one:List>";
        var xml = $"""<one:Page xmlns:one="{Ns}" name="P"><one:Title/>{Outline($"<one:OE>{bullet}<one:T>a</one:T></one:OE>")}{Outline("<one:OE><one:T>para</one:T></one:OE>")}{Outline($"<one:OE>{bullet}<one:T>b</one:T></one:OE>")}</one:Page>""";

        Assert.Equal("# P\n\n- a\n\npara\n\n- b", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void TreatsAnEmptyHeadingAsNoTitle() =>
        Assert.Null(MarkdownToOneNote.SplitTitle("#\n\nBody").Title);

    [Fact]
    public void EscapesBracketsInLinkText() =>
        Assert.Equal("# P\n\n[a\\]b \\[c](https://x)", Read("<one:OE><one:T><![CDATA[<a href=\"https://x\">a]b [c</a>]]></one:T></one:OE>"));

    [Fact]
    public void TitleKeepsCodeEntitiesAndLineBreaks()
    {
        Assert.Equal("Running dotnet test", MarkdownToOneNote.SplitTitle("# Running `dotnet test`\n\nBody").Title);
        Assert.Equal("Q&A", MarkdownToOneNote.SplitTitle("# Q&amp;A").Title);
        Assert.Equal("Release notes", MarkdownToOneNote.SplitTitle("Release\nnotes\n===\n\nBody").Title);
    }

    [Fact]
    public void DropsBoldThatCoversAWholeHeaderCell() =>
        Assert.Equal("# Sample\n\n| a and b | *c* |\n| --- | --- |\n| 1 | 2 |", RoundTrip("| **a** and **b** | *c* |\n| --- | --- |\n| 1 | 2 |"));

    [Fact]
    public void WrapsLinkDestinationsWithSpaces()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="L"><one:Title/>
              <one:Outline><one:OEChildren>
                <one:OE><one:T><![CDATA[See <a href="onenote:///C:/My Notes/a.one#x">Page (old)</a> and <a href="https://x.test/a_(b)">this</a>]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal("# L\n\nSee [Page (old)](<onenote:///C:/My Notes/a.one#x>) and [this](<https://x.test/a_(b)>)", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void TreatsOnlyMonospaceFontsAsCode()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="F"><one:Title/>
              <one:QuickStyleDef index="0" name="p" font="Cascadia Mono"/>
              <one:Outline><one:OEChildren>
                <one:OE quickStyleIndex="0"><one:T><![CDATA[Body in a monospace page font]]></one:T></one:OE>
                <one:OE style="font-family:'Monotype Corsiva'"><one:T><![CDATA[Script font]]></one:T></one:OE>
                <one:OE style="font-family:'Cascadia Code'"><one:T><![CDATA[Real code]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal("# F\n\nBody in a monospace page font\n\nScript font\n\n`Real code`", PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void KeepsImagesAsLinks() =>
        Assert.Equal("# Sample\n\n[Diagram](https://x.test/a.png)", RoundTrip("![Diagram](https://x.test/a.png)"));

    [Fact]
    public void DropsCodeLanguageButKeepsCode() =>
        Assert.Equal("# Sample\n\n```\nvar x = 1;\n```", RoundTrip("```csharp\nvar x = 1;\n```"));

    [Fact]
    public void EscapesTextTypedInOneNote()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="Typed"><one:Title/>
              <one:Outline><one:OEChildren>
                <one:OE><one:T><![CDATA[________________________]]></one:T></one:OE>
                <one:OE><one:T><![CDATA[# 1 priority: 2 * 3 = 6]]></one:T></one:OE>
                <one:OE><one:T><![CDATA[1. Not a real list]]></one:T></one:OE>
                <one:OE><one:List><one:Number numberSequence="0" numberFormat="##." text="a."/></one:List><one:T><![CDATA[Letter item]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal(@"# Typed

---

\# 1 priority: 2 \* 3 = 6

1\. Not a real list

1. Letter item".Replace("\r\n", "\n"), PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void SplitsLeadingTitle()
    {
        foreach (var markdown in new[] { "\n# My page\n\nBody", "﻿# My page\n\nBody", "#\tMy page\n\nBody", "  # My page #\n\nBody", "My page\n===\n\nBody" })
        {
            var (title, body) = MarkdownToOneNote.SplitTitle(markdown);
            Assert.Equal(("My page", "Body"), (title, body.Trim()));
        }

        Assert.Null(MarkdownToOneNote.SplitTitle("## Not a title\n# Later").Title);
    }

    [Fact]
    public void ReadsLooseListsAsItemText()
    {
        var md = RoundTrip("- One\n\n- Two\n\n  More about two");

        Assert.StartsWith("# Sample\n\n- One\n- Two\n\n  More about two", md);
    }

    // The way OneNote stores what HtmlToOneNote sends: spans for inline formatting, OE style for whole lines.
    [Fact]
    public void ReadsOneNoteStyledSpansAndCodeLines()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="Saved">
              <one:QuickStyleDef index="0" name="PageTitle" font="Aptos"/>
              <one:QuickStyleDef index="1" name="p" font="Aptos"/>
              <one:QuickStyleDef index="2" name="h2" font="Calibri"/>
              <one:QuickStyleDef index="3" name="code" font="Consolas"/>
              <one:Title><one:OE><one:T><![CDATA[Saved]]></one:T></one:OE></one:Title>
              <one:Outline><one:OEChildren>
                <one:OE quickStyleIndex="1"><one:T><![CDATA[<span style='font-family:Aptos'>A </span><span style='font-weight:bold;
            font-family:Aptos'>bold</span><span style='font-family:Aptos'> and </span><span style='font-family:Consolas'>code</span>]]></one:T></one:OE>
                <one:OE quickStyleIndex="2"><one:T><![CDATA[Heading]]></one:T></one:OE>
                <one:OE quickStyleIndex="3" style="font-family:Consolas;font-size:11.0pt"><one:T><![CDATA[if&nbsp;(x)]]></one:T></one:OE>
                <one:OE quickStyleIndex="3" style="font-family:Consolas;font-size:11.0pt"><one:T><![CDATA[&nbsp;&nbsp;y();]]></one:T></one:OE>
                <one:OE quickStyleIndex="1" style="font-family:Consolas;font-size:11.0pt"><one:T><![CDATA[npm install]]></one:T></one:OE>
                <one:OE quickStyleIndex="1" style="font-weight:bold"><one:T><![CDATA[Whole line bold]]></one:T></one:OE>
                <one:OE quickStyleIndex="1"><one:List><one:Number numberSequence="0" numberFormat="##." text="7."/></one:List><one:T><![CDATA[Seventh]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        // Only the Code style makes a code block. Other monospace text is inline code.
        Assert.Equal(
            "# Saved\n\nA **bold** and `code`\n\n## Heading\n\n```\nif (x)\n  y();\n```\n\n`npm install`\n\n**Whole line bold**\n\n7. Seventh",
            PageMarkdown.Convert(xml).Markdown);
    }

    [Fact]
    public void KeepsIndentedParagraphsOutsideListsFlush()
    {
        var xml = $"""
            <one:Page xmlns:one="{Ns}" name="P">
              <one:Outline><one:OEChildren>
                <one:OE><one:T><![CDATA[Parent]]></one:T><one:OEChildren>
                  <one:OE><one:T><![CDATA[Indented child]]></one:T></one:OE>
                </one:OEChildren></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal("# P\n\nParent\n\nIndented child", PageMarkdown.Convert(xml).Markdown);
    }
}
