using System.Text.RegularExpressions;

namespace OneNoteMcp;

/// <summary>
/// A page from a OneNote "Copy link to page" link.
/// </summary>
/// <param name="SectionPath">The section's full <c>.one</c> URL, or, for a web link, the site URL.</param>
/// <param name="RelativePath">For a web link, the section's path inside the site. Otherwise null.</param>
/// <param name="PageGuid">The <c>page-id</c> GUID.</param>
public sealed partial record PageLink(string SectionPath, string? RelativePath, string PageGuid)
{
    /// <summary>
    /// Parses the <c>onenote:</c> link or the <c>Doc.aspx</c> web link. Returns null for anything else.
    /// </summary>
    public static PageLink? Parse(string link)
    {
        // Match before decoding: the section path can contain encoded spaces.
        var m = OneNoteLink().Match(link);
        if (m.Success)
        {
            return new PageLink(Uri.UnescapeDataString(m.Groups[1].Value), null, m.Groups[2].Value);
        }

        m = WebLink().Match(Uri.UnescapeDataString(link));
        return m.Success ? new PageLink(m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value) : null;
    }

    /// <summary>
    /// True when a section's <c>path</c> attribute points to this link's section.
    /// </summary>
    public bool MatchesSection(string sectionPath)
    {
        var path = Uri.UnescapeDataString(sectionPath);
        return RelativePath is null
            ? path.Equals(SectionPath, StringComparison.OrdinalIgnoreCase)
            : path.StartsWith(SectionPath + "/", StringComparison.OrdinalIgnoreCase)
              && path.EndsWith("/" + RelativePath, StringComparison.OrdinalIgnoreCase);
    }

    // The page title after # can contain raw spaces.
    [GeneratedRegex(@"onenote:(https?://[^#\s]+?\.one)#[^\r\n]*?page-id=\{?([0-9A-Fa-f-]{36})")]
    private static partial Regex OneNoteLink();

    [GeneratedRegex(@"(https?://\S+?)/_layouts/\S*?wd=target\((.+?\.one)\|[0-9A-Fa-f-]{36}/.*\|([0-9A-Fa-f-]{36})/\)")]
    private static partial Regex WebLink();
}
