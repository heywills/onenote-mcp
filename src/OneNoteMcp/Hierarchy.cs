using System.Xml.Linq;

namespace OneNoteMcp;

public sealed record PageSummary(string Id, string Name, string? LastModified, int Level);

public sealed record SectionInfo(string Id, string Name, string? Path, IReadOnlyList<PageSummary> Pages);

public sealed record SectionGroupInfo(
    string Id, string Name, IReadOnlyList<SectionInfo> Sections, IReadOnlyList<SectionGroupInfo> SectionGroups);

public sealed record NotebookInfo(
    string Id, string Name, string? Path, string? LastModified,
    IReadOnlyList<SectionInfo> Sections, IReadOnlyList<SectionGroupInfo> SectionGroups);

public sealed record SearchResult(string PageId, string PageName, string Notebook, string Section, string LastModified);

/// <summary>
/// Parses OneNote hierarchy XML (from <c>GetHierarchy</c> and <c>FindPages</c>).
/// </summary>
public static class Hierarchy
{
    public static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    public static IReadOnlyList<NotebookInfo> ParseNotebooks(string xml) =>
        AsNotebooks(XElement.Parse(xml)).Elements(One + "Notebook").Select(nb => new NotebookInfo(
            Attr(nb, "ID"), Attr(nb, "name"), (string?)nb.Attribute("path"), (string?)nb.Attribute("lastModifiedTime"),
            ParseSections(nb), ParseSectionGroups(nb))).ToList();

    public static IReadOnlyList<SearchResult> ParseSearchResults(string xml) =>
        AsNotebooks(XElement.Parse(xml)).Elements(One + "Notebook")
            .SelectMany(nb => nb.Descendants(One + "Section").SelectMany(sec => sec.Elements(One + "Page").Select(p =>
                new SearchResult(Attr(p, "ID"), Attr(p, "name"), Attr(nb, "name"), Attr(sec, "name"), Attr(p, "lastModifiedTime")))))
            .ToList();

    /// <summary>
    /// Wraps a scoped result in Notebooks &gt; Notebook.
    /// </summary>
    /// <remarks>
    /// A notebook-, section-, or group-scoped <c>GetHierarchy</c> returns that node as the root.
    /// </remarks>
    public static XElement AsNotebooks(XElement root)
    {
        var tag = root.Name.LocalName;
        if (tag == "Notebooks")
        {
            return root;
        }

        return tag == "Notebook"
            ? new XElement(One + "Notebooks", root)
            : new XElement(One + "Notebooks", new XElement(One + "Notebook", root));
    }

    private static List<SectionInfo> ParseSections(XElement parent) =>
        parent.Elements(One + "Section").Select(sec => new SectionInfo(
            Attr(sec, "ID"), Attr(sec, "name"), (string?)sec.Attribute("path"),
            sec.Elements(One + "Page").Select(p => new PageSummary(
                Attr(p, "ID"), Attr(p, "name"), (string?)p.Attribute("lastModifiedTime"),
                int.TryParse((string?)p.Attribute("pageLevel"), out var level) ? level : 0)).ToList())).ToList();

    private static List<SectionGroupInfo> ParseSectionGroups(XElement parent) =>
        parent.Elements(One + "SectionGroup")
            .Where(sg => (string?)sg.Attribute("isRecycleBin") != "true")
            .Select(sg => new SectionGroupInfo(Attr(sg, "ID"), Attr(sg, "name"), ParseSections(sg), ParseSectionGroups(sg)))
            .ToList();

    private static string Attr(XElement e, string name) => (string?)e.Attribute(name) ?? "";
}
