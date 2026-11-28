using System.Xml.Linq;

namespace OneNoteMcp.Tests;

/// <summary>
/// OneNote tags in Markdown and HTML, both ways.
/// </summary>
public class TagTests
{
    private static readonly XNamespace One = Hierarchy.One;

    /// <summary>A page with To Do, Important, and Attendance tags on plain paragraphs and list items.</summary>
    private static readonly string TaggedPage = $"""
        <one:Page xmlns:one="{One}" name="P">
          <one:TagDef index="0" type="0" symbol="3" name="To Do"/>
          <one:TagDef index="1" type="1" symbol="13" name="Important"/>
          <one:TagDef index="2" type="106" symbol="94" name="Attendance"/>
          <one:Title/>
          <one:Outline><one:OEChildren>
            <one:OE><one:Tag index="0" completed="true"/><one:T><![CDATA[Call Ann]]></one:T></one:OE>
            <one:OE><one:Tag index="1" completed="false"/><one:T><![CDATA[Ship by Friday]]></one:T></one:OE>
            <one:OE><one:T><![CDATA[Plain]]></one:T></one:OE>
            <one:OE><one:Tag index="0" completed="false"/><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[Draft]]></one:T></one:OE>
            <one:OE><one:Tag index="2" completed="true"/><one:T><![CDATA[Mike]]></one:T></one:OE>
          </one:OEChildren></one:Outline>
        </one:Page>
        """;

    [Fact]
    public void WritesCheckboxesAsTaskItemsAndOtherTagsAsLabels() =>
        Assert.Equal("# P\n\n- [x] Call Ann\n\n[Important] Ship by Friday\n\nPlain\n\n- [ ] Draft\n- [x] Mike", PageMarkdown.Convert(TaggedPage).Markdown);

    [Fact]
    public void KeepsTaskItemsApartAndIndentsTheirChildren()
    {
        var page = $"""
            <one:Page xmlns:one="{One}" name="P"><one:TagDef index="0" type="0" symbol="3" name="To Do"/>
              <one:QuickStyleDef index="1" name="h2"/><one:Title/>
              <one:Outline><one:OEChildren>
                <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T>Bullet</one:T></one:OE>
                <one:OE><one:T></one:T></one:OE>
                <one:OE><one:Tag index="0" completed="false"/><one:T>Task</one:T>
                  <one:OEChildren><one:OE><one:T>Detail</one:T></one:OE></one:OEChildren></one:OE>
                <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T>[x] reviewed</one:T></one:OE>
                <one:OE quickStyleIndex="1"><one:Tag index="0" completed="true"/><one:T>Heading</one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        // The blank line keeps the bullet and the task in separate lists, the detail stays under its task,
        // untagged "[x]" text stays text, and a heading drops its checkbox
        Assert.Equal("# P\n\n- Bullet\n\n* [ ] Task\n\n  Detail\n\n* \\[x] reviewed\n\n## Heading", PageMarkdown.Convert(page).Markdown);
    }

    [Fact]
    public void WritesASecondCheckboxAsALabel()
    {
        Assert.Equal(("[ ] ", "[Discuss with] "), PageTags.Markdown([Tag.ToDo(false), new("Discuss with", 20, 94, true)]));
        Assert.Equal(("[x] ", ""), PageTags.Markdown([Tag.ToDo(false), Tag.ToDo(true)]));
    }

    [Fact]
    public void KeepsEmptyToDosAndNestedCheckboxText()
    {
        var page = $"""
            <one:Page xmlns:one="{One}" name="P"><one:TagDef index="0" type="0" symbol="3" name="To Do"/>
              <one:TagDef index="1" type="1" symbol="13" name="Important"/><one:Title/>
              <one:Outline><one:OEChildren>
                <one:OE><one:Tag index="0" completed="false"/><one:Tag index="1" completed="false"/><one:T></one:T></one:OE>
                <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T>a</one:T>
                  <one:OEChildren><one:OE><one:T>[ ] child</one:T></one:OE></one:OEChildren></one:OE>
                <one:OE><one:Tag index="0" completed="true"/><one:List><one:Bullet bullet="2"/></one:List>
                  <one:Image><one:CallbackID callbackID="cb"/></one:Image></one:OE>
                <one:OE><one:Tag index="0" completed="false"/><one:T></one:T></one:OE>
                <one:OE><one:T>Next</one:T></one:OE>
              </one:OEChildren></one:Outline>
            </one:Page>
            """;

        var markdown = PageMarkdown.Convert(page).Markdown;

        Assert.Equal("# P\n\n- [ ] [Important]\n- a\n\n  \\[ ] child\n\n- [x]\n\n  [Image 1]\n\n- [ ]\n\nNext", markdown);
        // And it reads back with the same tags, and "Next" stays its own paragraph
        var created = Create(MarkdownToOneNote.ToHtml(markdown[5..]), fromMarkdown: true);
        Assert.Equal(["false", "true", "false"], created.Descendants(One + "Tag").Select(t => t.Attribute("completed")!.Value));
        Assert.Contains(created.Descendants(One + "OE"), o => o.Element(One + "Tag") is null && o.Element(One + "T")?.Value == "Next");
    }

    [Fact]
    public void KeepsLinksLabelledX()
    {
        var page = $"""
            <one:Page xmlns:one="{One}" name="P"><one:TagDef index="0" type="0" symbol="3" name="To Do"/><one:Title/><one:Outline><one:OEChildren>
              <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[<a href="https://x.test/">x</a> marks it]]></one:T></one:OE>
              <one:OE><one:Tag index="0" completed="false"/><one:T><![CDATA[<a href="https://z.test/">x</a> to do]]></one:T></one:OE>
              <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[<a href="https://y.test/"> </a> blank]]></one:T></one:OE>
            </one:OEChildren></one:Outline></one:Page>
            """;

        var markdown = PageMarkdown.Convert(page).Markdown;
        // A space label is trimmed to [](url), which isn't a checkbox
        Assert.Equal("# P\n\n- [&#120;](https://x.test/) marks it\n- [ ] [&#120;](https://z.test/) to do\n- [](https://y.test/) blank", markdown);
        var created = Create(MarkdownToOneNote.ToHtml(markdown[5..]), fromMarkdown: true);
        // One To Do, on the item that had it, and every link kept
        Assert.Equal("false", created.Descendants(One + "Tag").Single().Attribute("completed")!.Value);
        Assert.Equal(3, created.Descendants(One + "T").Count(t => t.Value.Contains("<a href=")));
    }

    [Fact]
    public void KeepsTheDoneStateOfAnyTagInHtml()
    {
        var html = PageTags.Html([new Tag("Client request", 7, 50, true)]);
        var page = Create($"<p>{html}Text</p>", fromMarkdown: false);

        Assert.Equal("true", page.Descendants(One + "Tag").Single().Attribute("completed")!.Value);
    }

    [Fact]
    public void IgnoresCheckboxesInOtherHtml() =>
        Assert.Empty(Create("<ul><li><input type=\"CHECKBOX\" checked> Pasted</li></ul>", fromMarkdown: false).Descendants(One + "Tag"));

    [Fact]
    public void AddsStylesAfterTagsOnAPageWithoutStyles()
    {
        var page = new XElement(One + "Page",
            new XElement(One + "TagDef", new XAttribute("index", "0"), new XAttribute("type", "1"), new XAttribute("symbol", "13"), new XAttribute("name", "Important")),
            new XElement(One + "Title"));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("## Heading\n\n- [ ] Task"), fromMarkdown: true);

        Assert.Equal(["TagDef", "TagDef", "QuickStyleDef", "Title"], page.Elements().Take(4).Select(e => e.Name.LocalName));
    }

    [Fact]
    public void WritesTagsAsCheckboxesAndLabelsInHtml()
    {
        var html = PageHtml.Convert(TaggedPage).Html;

        Assert.Contains("<p><input type=\"checkbox\" disabled checked data-tag=\"To Do\" data-tag-type=\"0\" data-tag-symbol=\"3\"> Call Ann</p>", html);
        Assert.Contains("<p><span class=\"tag\" data-tag=\"Important\" data-tag-type=\"1\" data-tag-symbol=\"13\">[Important]</span> Ship by Friday</p>", html);
        Assert.Contains("<li><input type=\"checkbox\" disabled data-tag=\"To Do\" data-tag-type=\"0\" data-tag-symbol=\"3\"> Draft</li>", html);
    }

    [Fact]
    public void ReadsTagsBackFromHtml()
    {
        var page = Create(PageHtml.Convert(TaggedPage).Html, fromMarkdown: false);

        Assert.Equal(["To Do:0:3", "Important:1:13", "Attendance:106:94"],
            page.Elements(One + "TagDef").Select(d => $"{d.Attribute("name")!.Value}:{d.Attribute("type")!.Value}:{d.Attribute("symbol")!.Value}"));
        Assert.Equal(
            [("Call Ann", "0", "true", false), ("Ship by Friday", "1", "false", false), ("Draft", "0", "false", true), ("Mike", "2", "true", false)],
            Tagged(page));
    }

    [Fact]
    public void ReadsMarkdownTaskItemsAsToDos()
    {
        var page = Create(MarkdownToOneNote.ToHtml("- [ ] Open\n- [x] Done\n\n1. [x] Numbered\n\n- Plain bullet"), fromMarkdown: true);

        Assert.Equal("To Do", Assert.Single(page.Elements(One + "TagDef")).Attribute("name")!.Value);
        // Bulleted task items lose their bullet, like OneNote's own To Dos. Numbered ones keep their number.
        Assert.Equal([("Open", "0", "false", false), ("Done", "0", "true", false), ("Numbered", "0", "true", true)], Tagged(page));
        Assert.NotNull(page.Descendants(One + "OE").Single(o => o.Element(One + "T")?.Value == "Plain bullet").Element(One + "List"));
    }

    [Fact]
    public void ReusesThePagesTagDefinitions()
    {
        var page = new XElement(One + "Page",
            new XElement(One + "TagDef", new XAttribute("index", "0"), new XAttribute("type", "1"), new XAttribute("symbol", "13"), new XAttribute("name", "Important")),
            new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "PageTitle")),
            new XElement(One + "Title"));
        HtmlToOneNote.AddToPage(page, MarkdownToOneNote.ToHtml("- [ ] Task"), fromMarkdown: true);

        // TagDefs stay first, before the QuickStyleDefs, and a new one takes the next index
        Assert.Equal(["TagDef", "TagDef", "QuickStyleDef"], page.Elements().Take(3).Select(e => e.Name.LocalName));
        Assert.Equal("1", page.Elements(One + "TagDef").Last().Attribute("index")!.Value);
    }

    /// <summary>Markdown as onenote_get_page writes it reads back the same.</summary>
    internal const string TaskMarkdown = "- [ ] Open\n- [x] Done\n\n1. [x] Numbered\n\n[Important] A label stays text";

    [Fact]
    public void RoundTripsTaskItems()
    {
        var page = Create(MarkdownToOneNote.ToHtml(TaskMarkdown), fromMarkdown: true);
        page.SetAttributeValue("name", "P");

        Assert.Equal($"# P\n\n{TaskMarkdown}", PageMarkdown.Convert(page.ToString()).Markdown);
    }

    private static XElement Create(string html, bool fromMarkdown)
    {
        var page = new XElement(One + "Page", new XElement(One + "Title"));
        HtmlToOneNote.AddToPage(page, html, fromMarkdown);
        return page;
    }

    /// <summary>Each tagged OE's text, tag index, done state, and whether it's a list item.</summary>
    private static List<(string, string, string, bool)> Tagged(XElement page) =>
        [.. page.Descendants(One + "OE").Where(o => o.Element(One + "Tag") is not null).Select(o => (
            o.Element(One + "T")!.Value,
            o.Element(One + "Tag")!.Attribute("index")!.Value,
            o.Element(One + "Tag")!.Attribute("completed")!.Value,
            o.Element(One + "List") is not null))];
}
