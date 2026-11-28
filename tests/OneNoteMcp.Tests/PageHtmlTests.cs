using System.Xml.Linq;

namespace OneNoteMcp.Tests;

public class PageHtmlTests
{
    private static readonly XNamespace One = Hierarchy.One;

    private static string Body(string oes, string tableAndMore = "") =>
        PageHtml.Convert($"""
            <one:Page xmlns:one="{One}" name="My page">
              <one:QuickStyleDef index="0" name="PageTitle"/>
              <one:QuickStyleDef index="1" name="p"/>
              <one:QuickStyleDef index="2" name="h2"/>
              <one:QuickStyleDef index="3" name="blockquote"/>
              <one:QuickStyleDef index="4" name="code"/>
              <one:Title><one:OE><one:T><![CDATA[My page]]></one:T></one:OE></one:Title>
              <one:Outline><one:OEChildren>{oes}</one:OEChildren></one:Outline>{tableAndMore}
            </one:Page>
            """).Html.Split("<body>\n")[1].Split("</body>")[0];

    [Fact]
    public void WritesADocumentWithTheTitle()
    {
        var html = PageHtml.Convert($"""<one:Page xmlns:one="{One}" name="Q&amp;A"><one:Title/></one:Page>""").Html;
        Assert.StartsWith("<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n<title>Q&amp;A</title>", html);
    }

    [Fact]
    public void KeepsParagraphStyleAlignmentAndInlineHtml() =>
        Assert.Equal(
            "<p style=\"font-family:Calibri;font-size:12.0pt;color:#14324D;text-align:center\">Big <span style=\"color:red\">red</span></p>\n",
            Body("""<one:OE quickStyleIndex="1" alignment="center" style="font-family:Calibri;font-size:12.0pt;&#xA;color:#14324D"><one:T><![CDATA[Big <span style='color:red'>red</span>]]></one:T></one:OE>"""));

    [Fact]
    public void GroupsListsQuotesAndCode()
    {
        const string bullet = "<one:List><one:Bullet bullet=\"2\"/></one:List>";
        var html = Body($"""
            <one:OE>{bullet}<one:T>a</one:T><one:OEChildren><one:OE><one:List><one:Number numberSequence="0" numberFormat="##." text="3."/></one:List><one:T>three</one:T></one:OE></one:OEChildren></one:OE>
            <one:OE>{bullet}<one:T>b</one:T></one:OE>
            <one:OE quickStyleIndex="3"><one:T>quoted</one:T></one:OE>
            <one:OE quickStyleIndex="3"><one:T>more</one:T></one:OE>
            <one:OE quickStyleIndex="4"><one:T>x &lt; 1</one:T></one:OE>
            <one:OE quickStyleIndex="4"><one:T>  y</one:T></one:OE>
            <one:OE quickStyleIndex="2"><one:T>Heading</one:T></one:OE>
            """);

        Assert.Equal("""
            <ul>
            <li>a<ol start="3">
            <li>three</li>
            </ol>
            </li>
            <li>b</li>
            </ul>
            <blockquote>
            <p>quoted</p>
            <p>more</p>
            </blockquote>
            <pre>
            x &lt; 1
              y</pre>
            <h2>Heading</h2>

            """.Replace("\r\n", "\n"), html);
    }

    [Fact]
    public void IndentsChildrenOfAParagraph() =>
        Assert.Equal("<p>Parent</p>\n<div class=\"indent\" style=\"margin-left:2em\">\n<p>Child</p>\n</div>\n",
            Body("<one:OE><one:T>Parent</one:T><one:OEChildren><one:OE><one:T>Child</one:T></one:OE></one:OEChildren></one:OE>"));

    [Fact]
    public void KeepsTableHeadersAndShading() =>
        Assert.Equal("<table>\n<tr><th>\n<p>H</p>\n</th></tr>\n<tr><td style=\"background-color:#DDEBF7\">\n<p>Shaded</p>\n</td></tr>\n</table>\n",
            Body("""
                <one:OE><one:Table hasHeaderRow="true">
                  <one:Row><one:Cell><one:OEChildren><one:OE><one:T>H</one:T></one:OE></one:OEChildren></one:Cell></one:Row>
                  <one:Row><one:Cell shadingColor="#DDEBF7"><one:OEChildren><one:OE><one:T>Shaded</one:T></one:OE></one:OEChildren></one:Cell></one:Row>
                </one:Table></one:OE>
                """));

    [Fact]
    public void WritesImagesWithTheirSize()
    {
        var html = Body("""<one:OE><one:Image><one:Size width="120.5" height="60"/><one:CallbackID callbackID="cb"/></one:Image></one:OE>""");
        Assert.Equal("<p><img data-callback-id=\"cb\" alt=\"Image 1\" style=\"width:120.5pt;height:60pt\"></p>\n", html);
    }

    [Theory]
    [InlineData("png", "<img src=\"My%20page_images/image-2.png\" alt=\"Image 2\">")]
    [InlineData("bin", "<a href=\"My%20page_images/image-2.bin\">Image 2: unsupported type</a>")]
    public void LinksSavedImages(string extension, string html) =>
        Assert.Equal(html, PageHtml.ImageLink(new ImageRef("cb", 2, null, null), "My page_images", extension));

    [Fact]
    public void KeepsOnlySafeInlineHtml() =>
        Assert.Equal(
            "<p><b>b</b>bad <a href=\"https://x.test/?a=1&amp;b=2\">ok</a> <span style=\"color:red\">s</span>&nbsp;é</p>\n",
            Body("""<one:OE><one:T><![CDATA[<b>b</b><script>x()</script><a href="java&#9;script:alert(1)">bad</a> <a href="https://x.test/?a=1&b=2" onclick="x()">ok</a> <span style="color:red;position:fixed" lang="en-US">s</span><img src=x onerror=y>&nbsp;é]]></one:T></one:OE>"""));

    [Fact]
    public void KeepsAppLinksAndDropsScriptLinks() =>
        Assert.Equal("<p><a href=\"tel:+15551234\">Call</a> d</p>\n",
            Body("""<one:OE><one:T><![CDATA[<a href="tel:+15551234">Call</a> <a href="data:text/html,x">d</a>]]></one:T></one:OE>"""));

    [Fact]
    public void ShowsContentUnderACodeLine() =>
        Assert.Equal("<pre>\n\nx</pre>\n<p><img data-callback-id=\"cb\" alt=\"Image 1\"></p>\n<pre>\ny</pre>\n",
            Body("""
                <one:OE quickStyleIndex="4"><one:T></one:T></one:OE>
                <one:OE quickStyleIndex="4"><one:T>x</one:T><one:Image><one:CallbackID callbackID="cb"/></one:Image></one:OE>
                <one:OE quickStyleIndex="4"><one:T>y</one:T></one:OE>
                """));

    [Fact]
    public void LeavesImageOnlyLinesOutOfCode() =>
        Assert.Equal("<pre>\nx</pre>\n<p><img data-callback-id=\"cb\" alt=\"Image 1\"></p>\n<pre>\ny</pre>\n",
            Body("""
                <one:OE quickStyleIndex="4"><one:T>x</one:T></one:OE>
                <one:OE quickStyleIndex="4"><one:Image><one:CallbackID callbackID="cb"/></one:Image></one:OE>
                <one:OE quickStyleIndex="4"><one:T>y</one:T></one:OE>
                """));

    [Fact]
    public void KeepsCodeLinesWithNoText() =>
        Assert.Equal("<pre>\nx\n\ny</pre>\n",
            Body("""<one:OE quickStyleIndex="4"><one:T>x</one:T></one:OE><one:OE quickStyleIndex="4"/><one:OE quickStyleIndex="4"><one:T>y</one:T></one:OE>"""));

    [Fact]
    public void WritesBlankLinesThatShow() =>
        Assert.Equal("<p>a</p>\n<p><br></p>\n<p>b</p>\n", Body("<one:OE><one:T>a</one:T></one:OE><one:OE><one:T></one:T></one:OE><one:OE><one:T>b</one:T></one:OE>"));

    [Fact]
    public void FiltersParagraphStylesAndShading() =>
        Assert.Equal("<p style=\"color:red\">P</p>\n<table>\n<tr><td>\n<p>C</p>\n</td></tr>\n</table>\n",
            Body("""
                <one:OE style="position:fixed;color:red;background:url(https://x.test/a.png)"><one:T>P</one:T></one:OE>
                <one:OE><one:Table><one:Row><one:Cell shadingColor="red;background:url(x)"><one:OEChildren><one:OE><one:T>C</one:T></one:OE></one:OEChildren></one:Cell></one:Row></one:Table></one:OE>
                """));

    [Fact]
    public void PlacesSideBySideOutlines()
    {
        var html = PageHtml.Convert($"""
            <one:Page xmlns:one="{One}" name="P"><one:Title/>
              <one:Outline><one:Position x="36.0" y="86.4000015258789" z="0"/><one:Size width="540.0" height="459" isSetByUser="true"/>
                <one:OEChildren><one:OE><one:T>Main</one:T></one:OE></one:OEChildren></one:Outline>
              <one:Outline><one:Position x="612.0" y="104.400001525879" z="1"/><one:Size width="342.9000549316406" height="20"/>
                <one:OEChildren><one:OE><one:T>Note</one:T></one:OE></one:OEChildren></one:Outline>
              <one:Image><one:Position x="100" y="500" z="2"/><one:Size width="50" height="40"/><one:CallbackID callbackID="cb"/></one:Image>
            </one:Page>
            """).Html;

        // Only a width the user set is fixed. A page-level image keeps its place too.
        Assert.Contains("""
            <div class="outline" style="position:absolute;left:36pt;top:86.4pt;width:540pt;z-index:0">
            <p>Main</p>
            </div>
            <div class="outline" style="position:absolute;left:612pt;top:104.4pt;z-index:1">
            <p>Note</p>
            </div>
            <div class="outline" style="position:absolute;left:100pt;top:500pt;z-index:2">
            <p><img data-callback-id="cb" alt="Image 1" style="width:50pt;height:40pt"></p>
            </div>
            """, html);
    }

    [Fact]
    public void FlowsAPageWithOneOutlineOfText()
    {
        // An empty outline, and a page-level image without an ID, don't make it a placed page
        var html = PageHtml.Convert($"""
            <one:Page xmlns:one="{One}" name="P"><one:Title/>
              <one:Outline><one:Position x="36" y="86" z="0"/><one:OEChildren><one:OE><one:T>Text</one:T></one:OE></one:OEChildren></one:Outline>
              <one:Outline><one:Position x="400" y="86" z="1"/></one:Outline>
              <one:Image><one:Position x="100" y="500" z="2"/></one:Image>
            </one:Page>
            """).Html;

        Assert.DoesNotContain("class=\"outline\"", html);
        Assert.Contains("<p>Text</p>", html);
    }

    [Fact]
    public void DescribesImagesItCouldntSave() =>
        Assert.Equal("[Image 3: Can&#39;t save &lt;x&gt;]", PageHtml.ImageProblem(3, "Can't save <x>"));
}
