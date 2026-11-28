using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OneNoteMcp;

public sealed record PageSummary(string Id, string Name, string? LastModified, int Level);

public sealed record SectionInfo(string Id, string Name, string? Path, IReadOnlyList<PageSummary> Pages);

public sealed record SectionGroupInfo(
    string Id, string Name, IReadOnlyList<SectionInfo> Sections, IReadOnlyList<SectionGroupInfo> SectionGroups);

public sealed record NotebookInfo(
    string Id, string Name, string? Path, string? LastModified,
    IReadOnlyList<SectionInfo> Sections, IReadOnlyList<SectionGroupInfo> SectionGroups);

/// <summary>A page that matches a search.</summary>
/// <param name="InTitle">Every query word is in the page title</param>
public sealed record SearchResult(string PageId, string PageName, string Notebook, string Section, string LastModified, bool InTitle = false);

/// <summary>
/// Parses OneNote hierarchy XML (from <c>GetHierarchy</c> and <c>FindPages</c>).
/// </summary>
public static partial class Hierarchy
{
    public static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    public static IReadOnlyList<NotebookInfo> ParseNotebooks(string xml) =>
        AsNotebooks(XElement.Parse(xml)).Elements(One + "Notebook").Select(nb => new NotebookInfo(
            Attr(nb, "ID"), Attr(nb, "name"), (string?)nb.Attribute("path"), (string?)nb.Attribute("lastModifiedTime"),
            ParseSections(nb), ParseSectionGroups(nb))).ToList();

    /// <summary>The pages in hierarchy or <c>FindPages</c> XML, without those in Deleted Pages or the recycle bin.</summary>
    public static IReadOnlyList<SearchResult> ParseSearchResults(string xml) =>
        AsNotebooks(XElement.Parse(xml)).Elements(One + "Notebook")
            .SelectMany(nb => nb.Descendants(One + "Section").Where(sec => !IsDeleted(sec)).SelectMany(sec => sec.Elements(One + "Page")
                .Where(p => !IsTrue(p, "isInRecycleBin"))
                .Select(p => new SearchResult(Attr(p, "ID"), Attr(p, "name"), Attr(nb, "name"), Attr(sec, "name"), Attr(p, "lastModifiedTime")))))
            .ToList();

    /// <summary>
    /// Title matches first, then text matches from Windows Search, each newest first.
    /// </summary>
    /// <param name="query">The Windows Search query</param>
    /// <param name="pages">Every page in the search scope, from the hierarchy. Null for text matches only.</param>
    /// <param name="found">Windows Search's results</param>
    /// <remarks>
    /// - Title search finds pages the index misses, such as new pages.
    /// - A title matches when it contains every word from <see cref="TitleWords"/>, ignoring case. "art" matches "Party".
    /// </remarks>
    public static IReadOnlyList<SearchResult> RankSearchResults(string query, IEnumerable<SearchResult>? pages, IEnumerable<SearchResult> found)
    {
        var words = TitleWords(query);
        var titles = (words is null ? [] : pages ?? [])
            .Where(p => words!.All(w => p.PageName.Contains(w, StringComparison.OrdinalIgnoreCase)))
            .Select(p => p with { InTitle = true })
            .OrderByDescending(p => p.LastModified, StringComparer.Ordinal)
            .ToList();
        var text = found.OrderByDescending(p => p.LastModified, StringComparer.Ordinal);
        return [.. titles.Concat(text).DistinctBy(p => p.PageId)];
    }

    /// <summary>The words a title must contain, without quotes, parentheses, a trailing <c>*</c>, <c>AND</c>, or a lone dash. Null when titles aren't searched.</summary>
    /// <remarks>
    /// - A query with <c>OR</c>, <c>NOT</c>, a <c>-word</c>, a <c>+word</c>, or a property filter such as <c>title:plan</c> is text search only. Matching those in titles would need a parser for the Windows Search syntax.
    /// - A lone dash and parentheses are dropped, so a copied title such as "Nick Bushnell - 1:1 (draft)" still matches.
    /// </remarks>
    public static IReadOnlyList<string>? TitleWords(string query)
    {
        var words = query.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (words.Any(IsSearchSyntax))
        {
            return null;
        }

        var cleaned = words.Where(w => w != "AND" && w.Trim('-').Length > 0)
            .Select(w => w.Trim('"', '(', ')').TrimEnd('*').Trim('"', '(', ')'))
            .Where(w => w.Length > 0)
            .ToList();
        // No words left would match every title
        return cleaned.Count > 0 ? cleaned : null;
    }

    private static bool IsSearchSyntax(string word) =>
        word is "OR" or "NOT" || word.StartsWith('+') || (word.StartsWith('-') && word.Trim('-').Length > 0) || PropertyFilter().IsMatch(word);

    /// <summary>A Windows Search property filter, such as <c>author:x</c>. Not <c>Re:</c> or <c>1:1</c>.</summary>
    [GeneratedRegex(@"^[A-Za-z]+:\S")]
    private static partial Regex PropertyFilter();

    private static bool IsDeleted(XElement section) =>
        IsTrue(section, "isInRecycleBin") || IsTrue(section, "isDeletedPages")
        || section.Ancestors(One + "SectionGroup").Any(g => IsTrue(g, "isRecycleBin"));

    private static bool IsTrue(XElement e, string name) => (string?)e.Attribute(name) == "true";

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
