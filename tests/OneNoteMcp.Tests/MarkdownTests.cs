using System.Xml.Linq;

namespace OneNoteMcp.Tests;

public class MarkdownTests
{
    private const string Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private const string Sample = """
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

    [Fact]
    public void SplitsLeadingTitle()
    {
        Assert.Equal(("My page", "\nBody"), MarkdownToOneNote.SplitTitle("\n# My page\n\nBody"));
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
              <one:Outline><one:OEChildren>
                <one:OE quickStyleIndex="1"><one:T><![CDATA[<span style='font-family:Aptos'>A </span><span style='font-weight:bold;
            font-family:Aptos'>bold</span><span style='font-family:Aptos'> and </span><span style='font-family:Consolas'>code</span>]]></one:T></one:OE>
                <one:OE quickStyleIndex="2"><one:T><![CDATA[Heading]]></one:T></one:OE>
                <one:OE quickStyleIndex="1" style="font-family:Consolas;font-size:11.0pt"><one:T><![CDATA[if&nbsp;(x)]]></one:T></one:OE>
                <one:OE quickStyleIndex="1" style="font-family:Consolas;font-size:11.0pt"><one:T><![CDATA[&nbsp;&nbsp;y();]]></one:T></one:OE>
                <one:OE quickStyleIndex="1" style="font-weight:bold"><one:T><![CDATA[Whole line bold]]></one:T></one:OE>
                <one:OE quickStyleIndex="1"><one:List><one:Number numberSequence="0" numberFormat="##." text="7."/></one:List><one:T><![CDATA[Seventh]]></one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        Assert.Equal(
            "# Saved\n\nA **bold** and `code`\n\n## Heading\n\n```\nif (x)\n  y();\n```\n\n**Whole line bold**\n\n7. Seventh",
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
