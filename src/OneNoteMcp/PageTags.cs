using System.Net;
using System.Xml.Linq;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

/// <summary>A OneNote tag on a paragraph, such as To Do or Important.</summary>
/// <param name="Type">OneNote's tag type, kept so the tag reads back the same</param>
/// <param name="Symbol">OneNote's symbol for the tag</param>
/// <param name="Done">A checked checkbox</param>
public sealed record Tag(string Name, int Type, int Symbol, bool Done)
{
    /// <summary>The To Do tag that a Markdown task item becomes.</summary>
    public static Tag ToDo(bool done) => new("To Do", 0, 3, done);

    /// <summary>A checkbox tag, such as To Do or Attendance. Other tags, such as Important, are markers.</summary>
    /// <remarks>
    /// - The checkbox symbols seen in real pages: 3 (To Do) and 94 (Attendance, Discuss with). OneNote doesn't mark checkboxes otherwise.
    /// </remarks>
    public bool IsCheckbox => Symbol is 3 or 94;
}

/// <summary>
/// Reads and writes OneNote tags.
/// </summary>
/// <remarks>
/// - A page lists its tags as <c>TagDef</c>s. Each tagged OE holds <c>Tag</c>s that point to one by index.
/// - Markdown: a checkbox tag is a task item (<c>- [ ]</c>, <c>- [x]</c>). Other tags are their name in brackets, such as <c>[Important]</c>.
/// - HTML: a checkbox <c>input</c> or a <c>span class="tag"</c>, with <c>data-tag</c> attributes so every tag reads back.
/// </remarks>
public static class PageTags
{
    /// <summary>The page's tag definitions by index.</summary>
    public static Dictionary<string, Tag> Definitions(XElement page) =>
        page.Elements(One + "TagDef")
            .GroupBy(d => (string?)d.Attribute("index") ?? "")
            .ToDictionary(g => g.Key, g => new Tag(
                (string?)g.First().Attribute("name") ?? "Tag", ToInt(g.First().Attribute("type")), ToInt(g.First().Attribute("symbol")), false));

    /// <summary>The tags on an OE, in order.</summary>
    public static IReadOnlyList<Tag> Of(XElement oe, Dictionary<string, Tag> definitions) =>
        [.. oe.Elements(One + "Tag")
            .Where(t => definitions.ContainsKey((string?)t.Attribute("index") ?? ""))
            .Select(t => definitions[(string)t.Attribute("index")!] with { Done = (string?)t.Attribute("completed") == "true" })];

    /// <summary>Markdown for an OE's tags: the checkbox (<c>[ ] </c> or <c>[x] </c>), or null, and labels for the rest, such as <c>[Important] </c>.</summary>
    /// <remarks>
    /// - An item has one checkbox, so a different checkbox tag is a label too. A repeat of the same checkbox is dropped, and checks the box if it's done.
    /// </remarks>
    public static (string? Box, string Labels) Markdown(IReadOnlyList<Tag> tags)
    {
        var box = tags.FirstOrDefault(t => t.IsCheckbox) is { } first ? first with { Done = false } : null;
        bool SameBox(Tag t) => box is not null && t with { Done = false } == box;
        var labels = tags.Where(t => !SameBox(t));
        return (box is null ? null : tags.Any(t => SameBox(t) && t.Done) ? "[x] " : "[ ] ",
            string.Concat(labels.Select(t => $"[{PageMarkdown.EscapeBracketed(t.Name)}] ")));
    }

    /// <summary>HTML for an OE's tags, before its text.</summary>
    /// <remarks>
    /// - A label keeps its done state in <c>data-tag-done</c>, for a checkbox tag with a symbol not in <see cref="Tag.IsCheckbox"/>.
    /// </remarks>
    public static string Html(IReadOnlyList<Tag> tags) => string.Concat(tags.Select(t =>
    {
        var data = $" data-tag=\"{Encode(t.Name)}\" data-tag-type=\"{t.Type}\" data-tag-symbol=\"{t.Symbol}\"";
        return t.IsCheckbox
            ? $"<input type=\"checkbox\" disabled{(t.Done ? " checked" : "")}{data}> "
            : $"<span class=\"tag\"{data}{(t.Done ? " data-tag-done=\"true\"" : "")}>[{Encode(t.Name)}]</span> ";
    }));

    /// <summary>Adds the <c>Tag</c>s to an OE, first, where the schema puts them.</summary>
    /// <param name="index">The page's <c>TagDef</c> index for a tag, adding one when needed</param>
    public static void Add(XElement oe, IEnumerable<Tag> tags, Func<Tag, int> index)
    {
        foreach (var tag in tags.Reverse())
        {
            oe.AddFirst(new XElement(One + "Tag",
                new XAttribute("index", index(tag)), new XAttribute("completed", tag.Done ? "true" : "false"), new XAttribute("disabled", "false")));
        }
    }

    /// <summary>A <c>TagDef</c> for a tag.</summary>
    public static XElement Definition(Tag tag, int index) => new(One + "TagDef",
        new XAttribute("index", index), new XAttribute("type", tag.Type), new XAttribute("symbol", tag.Symbol),
        new XAttribute("fontColor", "automatic"), new XAttribute("highlightColor", "none"), new XAttribute("name", tag.Name));

    private static int ToInt(XAttribute? a) => int.TryParse((string?)a, out var n) ? n : 0;

    private static string Encode(string text) => WebUtility.HtmlEncode(text);
}
