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
/// - <c>pre</c>: one OE per line, in OneNote's Code style
/// - <c>blockquote</c>: paragraphs and lists in OneNote's Quote style. Headings, tables, and code inside a quote lose the quote.
/// </remarks>
public static partial class HtmlToOneNote
{
    // OneNote's built-in styles: (font, size, color, italic)
    private static readonly Dictionary<string, (string Font, string Size, string Color, bool Italic)> Styles = new()
    {
        ["h1"] = ("Calibri", "16.0", "#1E4E79", false),
        ["h2"] = ("Calibri", "14.0", "#2E75B5", false),
        ["h3"] = ("Calibri", "12.0", "#5B9BD5", false),
        ["h4"] = ("Calibri", "12.0", "#5B9BD5", true),
        ["h5"] = ("Calibri", "11.0", "#2E75B5", false),
        ["h6"] = ("Calibri", "11.0", "#2E75B5", true),
        // Not italic like OneNote's own Quote style: OneNote copies style italic onto the text, which hides italic words.
        ["blockquote"] = ("Calibri", "11.0", "#595959", false),
        ["code"] = ("Consolas", "10.0", "automatic", false),
    };

    private static readonly HashSet<string> Inline =
        ["b", "strong", "i", "em", "u", "s", "strike", "del", "sup", "sub", "a", "span", "code", "img"];

    /// <summary>
    /// Appends <paramref name="html"/> to a page as an Outline, and adds the styles it needs.
    /// </summary>
    public static void AddToPage(XElement page, string html)
    {
        var body = new HtmlParser().ParseDocument($"<body>{html}</body>").Body!;
        var defs = page.Elements(One + "QuickStyleDef").ToList();

        // Reuse the page's styles by name. New ones get the next free indexes.
        var indexes = defs.Where(d => d.Attribute("name") is not null)
            .GroupBy(d => (string)d.Attribute("name")!)
            .ToDictionary(g => g.Key, g => int.Parse((string?)g.First().Attribute("index") ?? "0"));
        var builder = new Builder(indexes, defs.Select(d => int.TryParse((string?)d.Attribute("index"), out var i) ? i : 0).DefaultIfEmpty(-1).Max() + 1);

        var outline = new XElement(One + "Outline");
        var children = new XElement(One + "OEChildren");
        outline.Add(children);
        builder.Blocks(body.ChildNodes, children);

        // QuickStyleDefs must sit together, before the content.
        var newDefs = builder.NewStyles.Select(style =>
        {
            var (font, size, color, italic) = Styles[style.Name];
            var def = new XElement(One + "QuickStyleDef",
                new XAttribute("index", style.Index), new XAttribute("name", style.Name),
                new XAttribute("fontColor", color), new XAttribute("highlightColor", "automatic"),
                new XAttribute("font", font), new XAttribute("fontSize", size),
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
    /// <param name="inCode">Keep every space, as <c>&amp;nbsp;</c></param>
    /// <param name="inLink">Inside a link, where an image can only be its alt text</param>
    private static string InlineHtml(IEnumerable<INode> nodes, bool inCode = false, bool inLink = false)
    {
        var sb = new StringBuilder();
        foreach (var n in nodes)
        {
            if (n is IText t)
            {
                sb.Append(inCode
                    ? WebUtility.HtmlEncode(t.Data).Replace(" ", "&nbsp;")
                    : WebUtility.HtmlEncode(Whitespace().Replace(t.Data, " ")));
                continue;
            }

            if (n is not IElement e)
            {
                continue;
            }

            var inner = InlineHtml(e.ChildNodes, inCode || e.LocalName == "code", inLink || e.LocalName == "a");
            var alt = e.GetAttribute("alt") is { Length: > 0 } a ? a : e.GetAttribute("src") ?? "image";
            sb.Append(e.LocalName switch
            {
                "b" or "strong" => $"<b>{inner}</b>",
                "i" or "em" => $"<i>{inner}</i>",
                "u" => $"<u>{inner}</u>",
                "s" or "strike" or "del" => $"<span style='text-decoration:line-through'>{inner}</span>",
                "sup" or "sub" => $"<{e.LocalName}>{inner}</{e.LocalName}>",
                "code" => $"<span style='font-family:Consolas'>{inner}</span>",
                "a" => $"<a href=\"{WebUtility.HtmlEncode(Href(e.GetAttribute("href")))}\">{inner}</a>",

                // OneNote can't load an image from a URL here: keep it as a link, or its alt text inside a link.
                "img" when inLink => WebUtility.HtmlEncode(alt),
                "img" => $"<a href=\"{WebUtility.HtmlEncode(Href(e.GetAttribute("src")))}\">{WebUtility.HtmlEncode(alt)}</a>",
                _ => inner,
            });
        }

        return sb.ToString();
    }

    /// <summary>
    /// The link target, with backslashes back.
    /// </summary>
    /// <remarks>Markdig writes a backslash in a URL as <c>%5C</c>, which breaks <c>\\server\share</c> links in OneNote.</remarks>
    private static string Href(string? href) => (href ?? "").Replace("%5C", "\\", StringComparison.OrdinalIgnoreCase);

    /// <param name="indexes">The page's QuickStyleDef indexes by name</param>
    /// <param name="nextIndex">The first free QuickStyleDef index</param>
    private sealed class Builder(Dictionary<string, int> indexes, int nextIndex)
    {
        private bool _inQuote;
        private int _listDepth;

        /// <summary>Per list depth: whether the next numbered list uses <c>##)</c>.</summary>
        /// <remarks>OneNote continues numbering across lists with the same format, even with paragraphs between them. Alternating formats restarts each list at 1.</remarks>
        private readonly Dictionary<int, bool> _altNumberFormat = [];

        /// <summary>Styles this content needs that the page doesn't have yet</summary>
        public List<(string Name, int Index)> NewStyles { get; } = [];

        private int StyleIndex(string name)
        {
            if (!indexes.TryGetValue(name, out var index))
            {
                index = indexes[name] = nextIndex++;
                NewStyles.Add((name, index));
            }

            return index;
        }

        /// <param name="number">For a numbered item, its number and format (<c>##.</c> or <c>##)</c>)</param>
        /// <remarks>numberSequence picks the number style (0 is 1, 2, 3; 1 is I, II, III), so it's always 0.</remarks>
        private static XElement Oe(XElement parent, string innerHtml, int? style = null, string? listKind = null, (int Value, string Format) number = default)
        {
            var oe = new XElement(One + "OE");
            if (style is not null)
            {
                oe.Add(new XAttribute("quickStyleIndex", style));
            }

            if (listKind is not null)
            {
                oe.Add(new XElement(One + "List", listKind == "ol"
                    ? new XElement(One + "Number", new XAttribute("numberSequence", "0"), new XAttribute("numberFormat", number.Format ?? "##."), new XAttribute("text", $"{number.Value}{(number.Format ?? "##.")[^1]}"))
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
                    Oe(parent, text, style: _inQuote ? StyleIndex("blockquote") : null);
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
                        if (parent.HasElements)
                        {
                            Oe(parent, ""); // blank line before a heading, like hand-made pages
                        }

                        Oe(parent, InlineHtml(e.ChildNodes), style: StyleIndex(e.LocalName));
                        break;
                    case "blockquote":
                        if (!_inQuote && parent.Elements(One + "OE").LastOrDefault() is { } previous && indexes.TryGetValue("blockquote", out var quoteIndex)
                            && (string?)previous.Attribute("quickStyleIndex") == quoteIndex.ToString())
                        {
                            Oe(parent, ""); // keeps back-to-back quotes apart
                        }

                        var outerQuote = _inQuote;
                        _inQuote = true;
                        Blocks(e.ChildNodes, parent);
                        _inQuote = outerQuote;
                        break;
                    case "ul" or "ol":
                        // Keeps back-to-back lists of the same kind apart. Bullets then numbers already read as two lists.
                        var kind = e.LocalName == "ol" ? "Number" : "Bullet";
                        if (parent.Elements(One + "OE").LastOrDefault()?.Element(One + "List")?.Element(One + kind) is not null)
                        {
                            Oe(parent, "");
                        }

                        ListItems(e, parent);
                        break;
                    case "table":
                        // Tables can't hold the quote style
                        var tableInQuote = _inQuote;
                        _inQuote = false;
                        Table(e, parent);
                        _inQuote = tableInQuote;
                        break;
                    case "pre":
                        var code = StyleIndex("code");
                        if (parent.Elements(One + "OE").LastOrDefault() is { } last && (string?)last.Attribute("quickStyleIndex") == code.ToString())
                        {
                            Oe(parent, ""); // keeps back-to-back code blocks apart
                        }

                        // Only the final newline ends the code. Leading and inner blank lines are code.
                        var codeText = e.TextContent.EndsWith('\n') ? e.TextContent[..^1] : e.TextContent;
                        foreach (var line in codeText.Split('\n'))
                        {
                            var escaped = WebUtility.HtmlEncode(line.TrimEnd('\r')).Replace(" ", "&nbsp;");
                            Oe(parent, escaped.Length > 0 ? escaped : "&nbsp;", style: code);
                        }

                        break;
                    case "hr":
                        Oe(parent, new string('―', 20));
                        break;
                    default: // p, div, section, etc.
                        Blocks(e.ChildNodes, parent);
                        break;
                }
            }

            Flush();
        }

        private void ListItems(IElement list, XElement parent)
        {
            var number = int.TryParse(list.GetAttribute("start"), out var start) ? start : 1;
            var format = "##.";
            if (list.LocalName == "ol")
            {
                var alt = _altNumberFormat.GetValueOrDefault(_listDepth);
                _altNumberFormat[_listDepth] = !alt;
                format = alt ? "##)" : "##.";
            }

            _listDepth++;
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

                var oe = Oe(parent, InlineHtml(inline), style: _inQuote ? StyleIndex("blockquote") : null, listKind: list.LocalName, number: (number++, format));
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

            _listDepth--;
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
