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
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml(markdown));
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
    public void RoundTripsEdgeCases(string markdown) =>
        Assert.Equal($"# Sample\n\n{markdown}", RoundTrip(markdown));

    [Fact]
    public void GivesEachNumberedListItsOwnSequence()
    {
        var page = new XElement(XName.Get("Page", Ns));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("1. A\n2. B\n\nText\n\n1. C"));

        Assert.Equal(["0", "0", "1"], page.Descendants(XName.Get("Number", Ns)).Select(n => (string?)n.Attribute("numberSequence")));
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
