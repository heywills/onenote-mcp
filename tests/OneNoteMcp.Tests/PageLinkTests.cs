namespace OneNoteMcp.Tests;

public class PageLinkTests
{
    private const string Guid = "0f1e2d3c-4b5a-6978-8796-a5b4c3d2e1f0";
    private const string Section = "https://contoso.sharepoint.com/sites/Team/Shared Documents/Team Notebook/Projects/Plan.one";

    private const string OneNoteLink =
        "onenote:https://contoso.sharepoint.com/sites/Team/Shared%20Documents/Team%20Notebook/Projects/Plan.one#Kickoff&section-id={11111111-2222-3333-4444-555555555555}&page-id={" + Guid + "}&end";

    private const string WebLink =
        "https://contoso.sharepoint.com/sites/Team/_layouts/Doc.aspx?sourcedoc={aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}&wd=target%28Projects%2FPlan.one%7C11111111-2222-3333-4444-555555555555%2FKickoff%7C" + Guid + "%2F%29";

    [Fact]
    public void ParsesOneNoteLink()
    {
        var link = PageLink.Parse(OneNoteLink)!;

        Assert.Equal((Section, null, Guid), (link.SectionPath, link.RelativePath, link.PageGuid));
        Assert.True(link.MatchesSection(Section.Replace(" ", "%20")));
        Assert.False(link.MatchesSection(Section.Replace("Plan", "Other")));
    }

    [Fact]
    public void ParsesOneNoteLinkWithSpacesInTitle() =>
        Assert.Equal(Guid, PageLink.Parse(OneNoteLink.Replace("#Kickoff", "#Kickoff meeting notes"))!.PageGuid);

    [Fact]
    public void ParsesWebLink()
    {
        var link = PageLink.Parse(WebLink)!;

        Assert.Equal(("https://contoso.sharepoint.com/sites/Team", "Projects/Plan.one", Guid), (link.SectionPath, link.RelativePath, link.PageGuid));
        Assert.True(link.MatchesSection(Section));
    }

    [Fact]
    public void PrefersOneNoteLinkWhenBothArePasted() =>
        Assert.Null(PageLink.Parse($"{WebLink}\n{OneNoteLink}")!.RelativePath);

    [Fact]
    public void RejectsOtherText() => Assert.Null(PageLink.Parse("https://example.com/page"));
}
