using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

/// <summary>
/// Converts simple HTML into OneNote page XML.
/// </summary>
/// <remarks>
/// OneNote's <c>one:T</c> holds inline HTML only. Block elements become OneNote structure:
/// - <c>p</c>, <c>div</c>, <c>br</c>: one OE per paragraph or line
/// - <c>h1</c>-<c>h6</c>: OE with a heading QuickStyleDef, and a blank OE before it
/// - <c>ul</c>, <c>ol</c>, <c>li</c>: OE with <c>one:List</c>, nested lists in OEChildren
/// - <c>table</c>: <c>one:Table</c>
/// - <c>pre</c>: one monospace OE per line
/// </remarks>
public static partial class HtmlToOneNote
{
    // OneNote's built-in heading styles: (font size, color, italic)
    private static readonly Dictionary<int, (string Size, string Color, bool Italic)> HeadingStyles = new()
    {
        [1] = ("16.0", "#1E4E79", false),
        [2] = ("14.0", "#2E75B5", false),
        [3] = ("12.0", "#5B9BD5", false),
        [4] = ("12.0", "#5B9BD5", true),
        [5] = ("11.0", "#2E75B5", false),
        [6] = ("11.0", "#2E75B5", true),
    };

    private static readonly HashSet<string> Inline =
        ["b", "strong", "i", "em", "u", "s", "strike", "del", "sup", "sub", "a", "span", "code"];

    /// <summary>
    /// Appends <paramref name="html"/> to a page as an Outline, and adds the heading styles it needs.
    /// </summary>
    public static void AddToPage(XElement page, string html)
    {
        var body = new HtmlParser().ParseDocument($"<body>{html}</body>").Body!;
        var builder = new Builder();
        var outline = new XElement(One + "Outline");
        var children = new XElement(One + "OEChildren");
        outline.Add(children);
        builder.Blocks(body.ChildNodes, children);

        // QuickStyleDef index = heading level. It must sit with the other QuickStyleDefs.
        var defs = page.Elements(One + "QuickStyleDef").ToList();
        var existing = defs.Select(d => (string?)d.Attribute("index")).ToHashSet();
        var newDefs = builder.Headings.Order()
            .Where(level => !existing.Contains(level.ToString()))
            .Select(level =>
            {
                var (size, color, italic) = HeadingStyles[level];
                var def = new XElement(One + "QuickStyleDef",
                    new XAttribute("index", level), new XAttribute("name", $"h{level}"),
                    new XAttribute("fontColor", color), new XAttribute("highlightColor", "automatic"),
                    new XAttribute("font", "Calibri"), new XAttribute("fontSize", size),
                    new XAttribute("spaceBefore", "0.0"), new XAttribute("spaceAfter", "0.0"));
                if (italic)
                {
                    def.Add(new XAttribute("italic", "true"));
                }

                return def;
            }).ToList();

        if (defs.Count > 0)
        {
            defs[^1].AddAfterSelf(newDefs);
        }
        else
        {
            page.AddFirst(newDefs);
        }

        page.Add(outline);
    }

    private static bool IsInline(INode n) => n is IText || (n is IElement e && Inline.Contains(e.LocalName));

    /// <summary>
    /// Renders inline nodes as the HTML OneNote accepts inside <c>one:T</c>.
    /// </summary>
    private static string InlineHtml(IEnumerable<INode> nodes)
    {
        var sb = new StringBuilder();
        foreach (var n in nodes)
        {
            if (n is IText t)
            {
                sb.Append(WebUtility.HtmlEncode(Whitespace().Replace(t.Data, " ")));
                continue;
            }

            if (n is not IElement e)
            {
                continue;
            }

            var inner = InlineHtml(e.ChildNodes);
            sb.Append(e.LocalName switch
            {
                "b" or "strong" => $"<b>{inner}</b>",
                "i" or "em" => $"<i>{inner}</i>",
                "u" => $"<u>{inner}</u>",
                "s" or "strike" or "del" => $"<span style='text-decoration:line-through'>{inner}</span>",
                "sup" or "sub" => $"<{e.LocalName}>{inner}</{e.LocalName}>",
                "code" => $"<span style='font-family:Consolas'>{inner}</span>",
                "a" => $"<a href=\"{WebUtility.HtmlEncode(e.GetAttribute("href") ?? "")}\">{inner}</a>",
                _ => inner,
            });
        }

        return sb.ToString();
    }

    private sealed class Builder
    {
        public HashSet<int> Headings { get; } = [];

        private static XElement Oe(XElement parent, string innerHtml, int? style = null, string? listKind = null)
        {
            var oe = new XElement(One + "OE");
            if (style is not null)
            {
                oe.Add(new XAttribute("quickStyleIndex", style));
            }

            if (listKind is not null)
            {
                oe.Add(new XElement(One + "List", listKind == "ol"
                    ? new XElement(One + "Number", new XAttribute("numberSequence", "0"), new XAttribute("numberFormat", "##."))
                    : new XElement(One + "Bullet", new XAttribute("bullet", "2"))));
            }

            oe.Add(new XElement(One + "T", new XCData(innerHtml.Trim())));
            parent.Add(oe);
            return oe;
        }

        /// <summary>
        /// Emits the block content of <paramref name="nodes"/> as OEs under an OEChildren.
        /// </summary>
        public void Blocks(IEnumerable<INode> nodes, XElement parent)
        {
            var run = new List<INode>();

            void Flush()
            {
                var text = InlineHtml(run).Trim();
                if (text.Length > 0)
                {
                    Oe(parent, text);
                }

                run.Clear();
            }

            foreach (var n in nodes)
            {
                if (IsInline(n))
                {
                    run.Add(n);
                    continue;
                }

                if (n is not IElement e)
                {
                    continue;
                }

                Flush();
                switch (e.LocalName)
                {
                    case "br":
                        break;
                    case "h1" or "h2" or "h3" or "h4" or "h5" or "h6":
                        var level = e.LocalName[1] - '0';
                        Headings.Add(level);
                        if (parent.HasElements)
                        {
                            Oe(parent, ""); // blank line before a heading, like hand-made pages
                        }

                        Oe(parent, InlineHtml(e.ChildNodes), style: level);
                        break;
                    case "ul" or "ol":
                        ListItems(e, parent);
                        break;
                    case "table":
                        Table(e, parent);
                        break;
                    case "pre":
                        foreach (var line in e.TextContent.Trim('\n').Split('\n'))
                        {
                            var escaped = WebUtility.HtmlEncode(line.TrimEnd('\r')).Replace(" ", "&nbsp;");
                            Oe(parent, $"<span style='font-family:Consolas'>{(escaped.Length > 0 ? escaped : "&nbsp;")}</span>");
                        }

                        break;
                    case "hr":
                        Oe(parent, new string('―', 20));
                        break;
                    default: // p, div, blockquote, section, etc.
                        Blocks(e.ChildNodes, parent);
                        break;
                }
            }

            Flush();
        }

        private void ListItems(IElement list, XElement parent)
        {
            foreach (var li in list.Children.Where(c => c.LocalName == "li"))
            {
                var inline = li.ChildNodes.Where(IsInline).ToList();
                var nested = li.ChildNodes.Where(c => !IsInline(c)).ToList();

                // Loose lists wrap the item text in <p>: use the first paragraph as the item text.
                if (string.IsNullOrWhiteSpace(InlineHtml(inline)) && nested.FirstOrDefault() is IElement { LocalName: "p" } p)
                {
                    inline = [.. p.ChildNodes];
                    nested.RemoveAt(0);
                }

                var oe = Oe(parent, InlineHtml(inline), listKind: list.LocalName);
                if (nested.Count > 0)
                {
                    var children = new XElement(One + "OEChildren");
                    Blocks(nested, children);
                    if (children.HasElements)
                    {
                        oe.Add(children);
                    }
                }
            }
        }

        private void Table(IElement table, XElement parent)
        {
            var rows = table.QuerySelectorAll("tr")
                .Select(r => r.Children.Where(c => c.LocalName is "td" or "th").ToList())
                .ToList();
            if (rows.Count == 0)
            {
                return;
            }

            var columns = Math.Max(1, rows.Max(r => r.Count));
            var tbl = new XElement(One + "Table", new XAttribute("bordersVisible", "true"),
                new XElement(One + "Columns", Enumerable.Range(0, columns).Select(i =>
                    new XElement(One + "Column", new XAttribute("index", i), new XAttribute("width", "120")))));

            foreach (var row in rows)
            {
                var r = new XElement(One + "Row");
                for (var i = 0; i < columns; i++)
                {
                    var cellChildren = new XElement(One + "OEChildren");
                    if (i < row.Count)
                    {
                        if (row[i].LocalName == "th")
                        {
                            Oe(cellChildren, $"<b>{InlineHtml(row[i].ChildNodes).Trim()}</b>");
                        }
                        else
                        {
                            Blocks(row[i].ChildNodes, cellChildren);
                        }
                    }

                    if (!cellChildren.HasElements)
                    {
                        Oe(cellChildren, "");
                    }

                    r.Add(new XElement(One + "Cell", cellChildren));
                }

                tbl.Add(r);
            }

            parent.Add(new XElement(One + "OE", tbl));
        }
    }

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();
}
