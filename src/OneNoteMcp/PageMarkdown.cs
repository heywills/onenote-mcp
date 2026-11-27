using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

public sealed record ImageRef(string CallbackId, int Index, double? Width, double? Height);

/// <summary>
/// Converts OneNote page XML to Markdown plus image references.
/// </summary>
/// <remarks>
/// - Headings come from the OE's QuickStyleDef name (<c>h1</c>-<c>h6</c>).
/// - Lists nest by indenting under the parent item. Paragraphs OneNote indents outside a list stay flush, or Markdown would read them as code.
/// - Code lines are OEs or text in a monospace font. Consecutive ones become one fenced block.
/// - OneNote keeps inline formatting as styled spans, and whole-paragraph formatting on the OE's <c>style</c>.
/// </remarks>
public static partial class PageMarkdown
{
    private static readonly HtmlParser Parser = new();

    public static (string Markdown, IReadOnlyList<ImageRef> Images) Convert(string xml)
    {
        var root = XElement.Parse(xml);
        var title = (string?)root.Attribute("name") ?? (string?)root.Attribute("ID") ?? "Untitled";
        var styles = root.Elements(One + "QuickStyleDef").ToDictionary(
            q => (string?)q.Attribute("index") ?? "",
            q => (Name: (string?)q.Attribute("name") ?? "", Font: (string?)q.Attribute("font") ?? ""));

        var writer = new Writer(styles, title);
        foreach (var outline in root.Descendants(One + "Outline"))
        {
            foreach (var children in outline.Elements(One + "OEChildren"))
            {
                writer.Block(children, "");
            }
        }

        // Images outside outlines
        foreach (var img in root.Descendants(One + "Image").Where(i => i.Ancestors(One + "Outline").FirstOrDefault() is null))
        {
            writer.Image(img, "");
        }

        return ($"# {title}\n\n{writer}".Trim(), writer.Images);
    }

    /// <summary>
    /// Strips the inline HTML OneNote keeps inside text elements.
    /// </summary>
    public static string CleanText(string text) =>
        System.Net.WebUtility.HtmlDecode(Tags().Replace(text, "")).Replace(' ', ' ');

    private sealed class Writer(Dictionary<string, (string Name, string Font)> styles, string title)
    {
        private readonly List<string> _lines = [];
        private string _kind = "";
        private string _codeIndent = "";
        private bool _titleChecked;
        private bool _topLevelBullet;

        public List<ImageRef> Images { get; } = [];

        public void Block(XElement oeChildren, string indent)
        {
            var number = 0;
            foreach (var oe in oeChildren.Elements(One + "OE"))
            {
                number = oe.Element(One + "List")?.Element(One + "Number") is null ? 0 : number + 1;
                Oe(oe, indent, number);
            }
        }

        /// <param name="number">Position in a numbered list, for items without OneNote's number text</param>
        private void Oe(XElement oe, string indent, int number)
        {
            var style = styles.TryGetValue((string?)oe.Attribute("quickStyleIndex") ?? "", out var s) ? s : (Name: "", Font: "");
            var oeStyle = ((string?)oe.Attribute("style") ?? "") + ";" + (style.Font.Length > 0 ? $"font-family:{style.Font}" : "");
            var list = oe.Element(One + "List");
            var marker = list?.Element(One + "Number") is { } numberElement ? ((string?)numberElement.Attribute("text") ?? $"{number}.") + " "
                : list is not null ? "- " : null;

            foreach (var t in oe.Elements(One + "T"))
            {
                var (text, isCode) = InlineMarkdown(t.Value, oeStyle + ";" + (string?)t.Attribute("style"));
                if (marker is not null)
                {
                    ListItem(indent, marker, text);
                }
                else if (isCode)
                {
                    CodeLine(indent, text);
                }
                else if (!string.IsNullOrWhiteSpace(text))
                {
                    var level = Regex.Match(style.Name, "^h([1-6])$");
                    Paragraph(indent, level.Success ? $"{new string('#', int.Parse(level.Groups[1].Value))} {text.Trim()}" : text.Trim());
                }
            }

            foreach (var child in oe.Elements())
            {
                switch (child.Name.LocalName)
                {
                    case "Image":
                        Image(child, indent);
                        break;
                    case "Table":
                        Table(child, indent);
                        break;
                    case "InsertedFile":
                        Paragraph(indent, $"[Attached: {(string?)child.Attribute("preferredName") ?? "file"}]");
                        break;
                    case "OEChildren":
                        // Only list items indent their children. Indented text elsewhere would read as code.
                        Block(child, marker is null ? indent : indent + new string(' ', marker.Length));
                        break;
                }
            }
        }

        public void Image(XElement img, string indent)
        {
            var id = (string?)img.Element(One + "CallbackID")?.Attribute("callbackID") ?? (string?)img.Attribute("callbackID");
            if (id is null || Images.Any(i => i.CallbackId == id))
            {
                return;
            }

            var size = img.Element(One + "Size");
            var reference = new ImageRef(id, Images.Count + 1, ToDouble(size?.Attribute("width")), ToDouble(size?.Attribute("height")));
            Images.Add(reference);
            Paragraph(indent, $"[Image {reference.Index}]");
        }

        private void Table(XElement table, string indent)
        {
            var rows = table.Elements(One + "Row").Select((row, r) => row.Elements(One + "Cell").Select(cell =>
            {
                var text = string.Join(" ", cell.Descendants(One + "OE").SelectMany(oe => oe.Elements(One + "T").Select(t =>
                    InlineMarkdown(t.Value, r == 0 ? "" : (string?)oe.Attribute("style") ?? "").Text.Trim())));
                text = Whitespace().Replace(text, " ").Trim().Replace("|", "\\|");

                // The header row renders bold already
                return r == 0 && text.Length > 4 && text.StartsWith("**") && text.EndsWith("**") ? text[2..^2] : text;
            }).ToList()).ToList();
            if (rows.Count == 0)
            {
                return;
            }

            Separate();
            var columns = rows.Max(r => r.Count);
            foreach (var (row, i) in rows.Select((row, i) => (row, i)))
            {
                _lines.Add($"{indent}| {string.Join(" | ", row.Concat(Enumerable.Repeat("", columns - row.Count)))} |");
                if (i == 0)
                {
                    _lines.Add($"{indent}| {string.Join(" | ", Enumerable.Repeat("---", columns))} |");
                }
            }

            _kind = "table";
            foreach (var img in table.Descendants(One + "Image"))
            {
                Image(img, indent);
            }
        }

        private void Paragraph(string indent, string text)
        {
            // The page body often repeats the title on its first line.
            if (!_titleChecked)
            {
                _titleChecked = true;
                if (text.Trim() == title.Trim())
                {
                    return;
                }
            }

            Separate();
            _lines.Add(indent + text);
            _kind = "paragraph";
        }

        private void ListItem(string indent, string marker, string text)
        {
            _titleChecked = true;
            var bullet = marker == "- ";

            // A new list, or a top-level switch between bullets and numbers
            if (_kind != "list" || (indent.Length == 0 && bullet != _topLevelBullet))
            {
                Separate();
            }

            if (indent.Length == 0)
            {
                _topLevelBullet = bullet;
            }

            _lines.Add(indent + marker + text.Trim());
            _kind = "list";
        }

        private void CodeLine(string indent, string text)
        {
            if (text.Length == 0 && _kind != "code")
            {
                return; // a blank monospace line outside code
            }

            _titleChecked = true;
            if (_kind != "code")
            {
                Separate();
                _codeIndent = indent;
                _lines.Add(indent + "```");
            }

            _lines.Add(indent + text.TrimEnd());
            _kind = "code";
        }

        /// <summary>
        /// Closes an open code block and adds a blank line before the next block.
        /// </summary>
        private void Separate()
        {
            if (_kind == "code")
            {
                _lines.Add(_codeIndent + "```");
            }

            if (_lines.Count > 0 && _lines[^1].Length > 0)
            {
                _lines.Add("");
            }

            _kind = "";
        }

        public override string ToString()
        {
            Separate();
            return string.Join("\n", _lines).Trim();
        }
    }

    /// <summary>
    /// Converts the HTML inside a text element to Markdown.
    /// </summary>
    /// <returns>The Markdown, and whether all of its text is monospace (a code line).</returns>
    private static (string Text, bool IsCode) InlineMarkdown(string html, string paragraphStyle)
    {
        var body = Parser.ParseDocument($"<body>{html}</body>").Body!;
        var paragraph = StyleOf(paragraphStyle);
        var state = new InlineState();
        var text = Inline(body.ChildNodes, paragraph.Code, state);

        if (!state.HasPlainText && !state.HasCodeText)
        {
            return ("", paragraph.Code || state.HasCodeSpace); // blank line, kept inside a code block
        }

        if (state.HasPlainText)
        {
            text = Wrap(text, paragraph.Strike ? "~~" : "");
            text = Wrap(text, paragraph.Italic ? "*" : "");
            return (Wrap(text, paragraph.Bold ? "**" : ""), false);
        }

        // A code line: plain text, spaces kept
        return (body.TextContent.Replace(' ', ' '), true);
    }

    private sealed class InlineState
    {
        public bool HasPlainText { get; set; }

        public bool HasCodeText { get; set; }

        /// <summary>Any monospace text, even just spaces: a blank line inside code.</summary>
        public bool HasCodeSpace { get; set; }
    }

    private static string Inline(INodeList nodes, bool inCode, InlineState state)
    {
        var sb = new StringBuilder();
        foreach (var node in nodes)
        {
            if (node is IText t)
            {
                var text = t.Data.Replace(' ', ' ');
                state.HasCodeSpace |= inCode;
                if (text.Trim().Length > 0)
                {
                    if (inCode)
                    {
                        state.HasCodeText = true;
                    }
                    else
                    {
                        state.HasPlainText = true;
                    }
                }

                sb.Append(Whitespace().Replace(text, " "));
                continue;
            }

            if (node is not IElement e)
            {
                continue;
            }

            var style = StyleOf(e.GetAttribute("style") ?? "");
            var tag = e.LocalName;
            var code = inCode || style.Code || tag == "code";
            var inner = Inline(e.ChildNodes, code, state);

            if (tag == "br")
            {
                sb.Append(' ');
                continue;
            }

            if (tag == "a")
            {
                sb.Append($"[{inner.Trim()}]({e.GetAttribute("href")})");
                continue;
            }

            if (code && !inCode)
            {
                inner = Wrap(inner, "`");
            }
            else if (!code)
            {
                inner = Wrap(inner, style.Strike || tag is "s" or "del" or "strike" ? "~~" : "");
                inner = Wrap(inner, style.Italic || tag is "i" or "em" ? "*" : "");
                inner = Wrap(inner, style.Bold || tag is "b" or "strong" ? "**" : "");
            }

            sb.Append(inner);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Wraps text in a Markdown marker, keeping edge spaces outside so the marker hugs the text.
    /// </summary>
    private static string Wrap(string text, string marker)
    {
        var core = text.Trim();
        if (marker.Length == 0 || core.Length == 0)
        {
            return text;
        }

        var lead = text[..(text.Length - text.TrimStart().Length)];
        var trail = text[text.TrimEnd().Length..];
        return $"{lead}{marker}{core}{marker}{trail}";
    }

    private static (bool Bold, bool Italic, bool Strike, bool Code) StyleOf(string css)
    {
        var s = css.ToLowerInvariant().Replace(" ", "").Replace("\n", "").Replace("\r", "");
        return (s.Contains("font-weight:bold"), s.Contains("font-style:italic"), s.Contains("line-through"), Monospace().IsMatch(s));
    }

    private static double? ToDouble(XAttribute? a) =>
        double.TryParse((string?)a, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : null;

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex Tags();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();

    [GeneratedRegex("font-family:[^;]*(consolas|courier|mono)")]
    private static partial Regex Monospace();
}
