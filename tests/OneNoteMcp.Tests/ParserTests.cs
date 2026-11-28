namespace OneNoteMcp.Tests;

public class ParserTests
{
    private const string Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private const string HierarchyXml = $"""
        <one:Notebooks xmlns:one="{Ns}">
          <one:Notebook name="Work Notes" ID="nb-001" path="C:\Users\test\Work Notes" lastModifiedTime="2026-02-15T10:00:00Z">
            <one:Section name="Meeting Notes" ID="sec-001" path="C:\Users\test\Work Notes\Meeting Notes.one">
              <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z" pageLevel="0"/>
              <one:Page ID="page-002" name="Sprint Review" lastModifiedTime="2026-02-13T14:00:00Z" pageLevel="2"/>
            </one:Section>
            <one:SectionGroup name="Archive" ID="sg-001">
              <one:Section name="Old Notes" ID="sec-002">
                <one:Page ID="page-003" name="Archived Page" pageLevel="0"/>
              </one:Section>
            </one:SectionGroup>
            <one:SectionGroup name="Recycle Bin" ID="sg-bin" isRecycleBin="true">
              <one:Section name="Deleted" ID="sec-deleted"/>
            </one:SectionGroup>
          </one:Notebook>
          <one:Notebook name="Personal" ID="nb-002" path="C:\Users\test\Personal">
            <one:Section name="Journal" ID="sec-003">
              <one:Page ID="page-004" name="Feb 10" pageLevel="0"/>
            </one:Section>
          </one:Notebook>
        </one:Notebooks>
        """;

    private const string PageXml = $"""
        <one:Page xmlns:one="{Ns}" ID="page-001" name="Test Page">
          <one:Title><one:OE><one:T><![CDATA[Test Page]]></one:T></one:OE></one:Title>
          <one:Outline>
            <one:OEChildren>
              <one:OE><one:T><![CDATA[This is paragraph one.]]></one:T></one:OE>
              <one:OE><one:T><![CDATA[This is paragraph two with <b>bold</b> text &amp; more.]]></one:T></one:OE>
              <one:OE>
                <one:Image>
                  <one:Size width="640" height="480" isSetByUser="true"/>
                  <one:CallbackID callbackID="img-001"/>
                </one:Image>
              </one:OE>
              <one:OE><one:T><![CDATA[Text after image.]]></one:T></one:OE>
            </one:OEChildren>
          </one:Outline>
          <one:Outline>
            <one:OEChildren>
              <one:OE>
                <one:Table>
                  <one:Row>
                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 1]]></one:T></one:OE></one:OEChildren></one:Cell>
                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 2]]></one:T></one:OE></one:OEChildren></one:Cell>
                  </one:Row>
                  <one:Row>
                    <one:Cell><one:OEChildren>
                      <one:OE><one:T><![CDATA[Data 1]]></one:T></one:OE>
                      <one:OE><one:T><![CDATA[second line]]></one:T></one:OE>
                    </one:OEChildren></one:Cell>
                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 2]]></one:T></one:OE></one:OEChildren></one:Cell>
                  </one:Row>
                </one:Table>
              </one:OE>
            </one:OEChildren>
          </one:Outline>
        </one:Page>
        """;

    [Fact]
    public void ParsesNotebooksSectionsAndPages()
    {
        var notebooks = Hierarchy.ParseNotebooks(HierarchyXml);

        Assert.Equal(["Work Notes", "Personal"], notebooks.Select(n => n.Name));
        var work = notebooks[0];
        Assert.Equal("nb-001", work.Id);
        Assert.Equal(["Monday Standup", "Sprint Review"], work.Sections.Single().Pages.Select(p => p.Name));
        Assert.Equal(2, work.Sections[0].Pages[1].Level);
    }

    [Fact]
    public void SkipsRecycleBinAndKeepsNestedGroups()
    {
        var group = Hierarchy.ParseNotebooks(HierarchyXml)[0].SectionGroups.Single();

        Assert.Equal("Archive", group.Name);
        Assert.Equal("Archived Page", group.Sections.Single().Pages.Single().Name);
    }

    [Theory]
    [InlineData("""<one:Notebook xmlns:one="{0}" ID="nb" name="N"><one:Section ID="s" name="S"/></one:Notebook>""")]
    [InlineData("""<one:Section xmlns:one="{0}" ID="s" name="S"/>""")]
    [InlineData("""<one:SectionGroup xmlns:one="{0}" ID="g" name="G"><one:Section ID="s" name="S"/></one:SectionGroup>""")]
    public void ParsesScopedResults(string template)
    {
        var notebooks = Hierarchy.ParseNotebooks(string.Format(template, Ns));

        var notebook = Assert.Single(notebooks);
        Assert.Contains(notebook.Sections.Concat(notebook.SectionGroups.SelectMany(g => g.Sections)), s => s.Id == "s");
    }

    [Fact]
    public void ConvertsPageText()
    {
        var (md, _) = PageMarkdown.Convert(PageXml);

        Assert.StartsWith("# Test Page\n\nThis is paragraph one.\n\nThis is paragraph two with **bold** text & more.", md);
        Assert.Contains("Text after image.", md);
    }

    [Fact]
    public void SkipsFirstLineOnlyWhenOneNoteNamedThePageFromIt()
    {
        var bodyStartsWithTitle = PageXml.Replace("This is paragraph one.", "Test Page]]></one:T></one:OE><one:OE><one:T><![CDATA[This is paragraph one.");
        var noTitleElement = System.Text.RegularExpressions.Regex.Replace(bodyStartsWithTitle, "<one:Title>.*?</one:Title>", "", System.Text.RegularExpressions.RegexOptions.Singleline);

        Assert.StartsWith("# Test Page\n\nThis is paragraph one.", PageMarkdown.Convert(noTitleElement).Markdown);
        Assert.StartsWith("# Test Page\n\nTest Page\n\nThis is paragraph one.", PageMarkdown.Convert(bodyStartsWithTitle).Markdown);
    }

    [Fact]
    public void ListsImages()
    {
        var (md, images) = PageMarkdown.Convert(PageXml);

        var image = Assert.Single(images);
        Assert.Equal(("img-001", 640.0, 480.0), (image.CallbackId, image.Width!.Value, image.Height!.Value));
        Assert.Contains("[Image 1]", md);
    }

    [Fact]
    public void ConvertsTableOnceWithCellLinesJoined()
    {
        var (md, _) = PageMarkdown.Convert(PageXml);

        Assert.Contains("| Header 1 | Header 2 |\n| --- | --- |\n| Data 1 second line | Data 2 |", md);
        Assert.Single(md.Split('\n'), l => l.Contains("Data 2"));
    }

    [Fact]
    public void ParsesSearchResults()
    {
        var result = Assert.Single(Hierarchy.ParseSearchResults($"""
            <one:Notebooks xmlns:one="{Ns}">
              <one:Notebook name="Work Notes" ID="nb-001">
                <one:SectionGroup name="G" ID="g"><one:Section name="Meeting Notes" ID="sec-001">
                  <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z"/>
                </one:Section></one:SectionGroup>
              </one:Notebook>
            </one:Notebooks>
            """));

        Assert.Equal(("Monday Standup", "Work Notes", "Meeting Notes"), (result.PageName, result.Notebook, result.Section));
        Assert.Empty(Hierarchy.ParseSearchResults($"""<one:Notebooks xmlns:one="{Ns}"/>"""));
    }

    [Fact]
    public void SkipsDeletedPagesInSearchResults() =>
        Assert.Equal(["kept"], Hierarchy.ParseSearchResults($"""
            <one:Notebook xmlns:one="{Ns}" name="N" ID="nb">
              <one:Section name="S" ID="s"><one:Page ID="kept" name="Kept"/><one:Page ID="p1" name="Deleted" isInRecycleBin="true"/></one:Section>
              <one:SectionGroup name="Recycle Bin" ID="bin" isRecycleBin="true">
                <one:Section name="Deleted Pages" ID="dp" isDeletedPages="true"><one:Page ID="p2" name="Deleted"/></one:Section>
                <one:Section name="Old" ID="old"><one:Page ID="p3" name="Deleted"/></one:Section>
              </one:SectionGroup>
              <one:Section name="Gone" ID="gone" isInRecycleBin="true"><one:Page ID="p4" name="Deleted"/></one:Section>
            </one:Notebook>
            """).Select(r => r.PageId));

    [Fact]
    public void RanksTitleMatchesFirstThenNewest()
    {
        static SearchResult Page(string id, string name, string modified) => new(id, name, "N", "S", modified);
        var pages = new[] { Page("t-old", "Nick Bushnell - 1:1", "2025"), Page("t-new", "1:1 with nick", "2026"), Page("other", "Nick notes", "2027") };
        var found = new[] { Page("x-old", "Standup", "2024"), Page("t-new", "1:1 with nick", "2026"), Page("x-new", "Retro", "2026-11") };

        var ranked = Hierarchy.RankSearchResults("nick 1:1", pages, found);

        Assert.Equal(["t-new", "t-old", "x-new", "x-old"], ranked.Select(r => r.PageId));
        Assert.Equal([true, true, false, false], ranked.Select(r => r.InTitle));
        Assert.Equal(["x-new", "t-new", "x-old"], Hierarchy.RankSearchResults("nick 1:1", null, found).Select(r => r.PageId));
    }

    [Theory]
    [InlineData("\"project plan\"", true)]
    [InlineData("proj* AND plan", true)]
    [InlineData("project NOT plan", false)]
    [InlineData("project NOT draft", false)]
    [InlineData("project -draft", false)]
    [InlineData("project OR plan", false)]
    [InlineData("(project) plan", true)]
    [InlineData("project - plan", true)]
    [InlineData("project +plan", false)]
    [InlineData("title:plan", false)]
    [InlineData("\"\" * AND -- ()", false)]
    public void MatchesTitlesWithoutSearchSyntax(string query, bool matches) =>
        Assert.Equal(matches, Hierarchy.RankSearchResults(query, [new("p", "Project plan", "N", "S", "2026")], []).Count == 1);

    [Theory]
    [InlineData("Nick Bushnell - 1:1")]
    [InlineData("Re: Live search (deleting now)")]
    public void MatchesCopiedTitles(string title) =>
        Assert.Single(Hierarchy.RankSearchResults(title, [new("p", title, "N", "S", "2026")], []));
}
