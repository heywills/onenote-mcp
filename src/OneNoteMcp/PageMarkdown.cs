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
/// - Block types come from the OE's QuickStyleDef name: <c>h1</c>-<c>h6</c>, <c>blockquote</c>, and <c>code</c> (fenced code).
/// - Lists nest by indenting under the parent item. Paragraphs OneNote indents outside a list stay flush, or Markdown would read them as code.
/// - Other monospace text becomes inline code.
/// - OneNote keeps inline formatting as styled spans, and whole-paragraph formatting on the OE's <c>style</c>.
/// - Plain text is escaped, so text that looks like Markdown stays text.
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
            q => (Name: (string?)q.Attribute("name") ?? "", Italic: (string?)q.Attribute("italic") == "true"));

        // Without a Title element, OneNote names the page after its first line, so that line repeats the name.
        var repeatedTitle = root.Element(One + "Title") is null
            ? root.Elements(One + "Outline").FirstOrDefault()?.Element(One + "OEChildren")?.Element(One + "OE")
            : null;
        if (repeatedTitle is not null && CleanText(string.Concat(repeatedTitle.Elements(One + "T").Select(t => t.Value))).Trim() != title.Trim())
        {
            repeatedTitle = null;
        }

        var writer = new Writer(styles, repeatedTitle);
        foreach (var outline in root.Descendants(One + "Outline"))
        {
            writer.NewOutline();
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

        return ($"# {EscapeClosingHashes(Escape(title))}\n\n{writer}".Trim(), writer.Images);
    }

    /// <summary>
    /// Strips the inline HTML OneNote keeps inside text elements.
    /// </summary>
    public static string CleanText(string text) =>
        System.Net.WebUtility.HtmlDecode(Tags().Replace(text, "")).Replace(' ', ' ');

    private sealed class Writer(Dictionary<string, (string Name, bool Italic)> styles, XElement? repeatedTitle)
    {
        private readonly List<string> _lines = [];
        private string _kind = "";
        private int _codeStart;
        private string _codeIndent = "";

        /// <summary>The line prefix of the last quoted block, such as <c>"> "</c> or <c>"  > "</c>. Null when it wasn't quoted.</summary>
        private string? _quotePrefix;
        private bool _topLevelBullet;

        /// <summary>Use <c>*</c> and <c>1)</c> markers.</summary>
        /// <remarks>Side-by-side outlines that both hold lists would read as one list.</remarks>
        private bool _altMarkers;

        public List<ImageRef> Images { get; } = [];

        public void NewOutline()
        {
            _altMarkers = _kind == "list" && !_altMarkers;

            Separate(null);
        }

        /// <param name="quoteColumn">The indent where an enclosing quote's <c>"> "</c> sits. Null outside a quote.</param>
        public void Block(XElement oeChildren, string indent, string? quoteColumn = null)
        {
            var number = 0;
            foreach (var oe in oeChildren.Elements(One + "OE"))
            {
                number = oe.Element(One + "List")?.Element(One + "Number") is null ? 0 : number + 1;
                Oe(oe, indent, number, quoteColumn);
            }
        }

        /// <param name="number">Position in a numbered list, for items without OneNote's number text</param>
        private void Oe(XElement oe, string indent, int number, string? quoteColumn)
        {
            var (styleName, styleItalic) = styles.TryGetValue((string?)oe.Attribute("quickStyleIndex") ?? "", out var s) ? s : ("", false);
            var quote = styleName == "blockquote";

            // Inside an enclosing quote, "> " stays at that quote's column and the rest of the indent follows it.
            quoteColumn = quote ? quoteColumn ?? indent : null;
            var prefix = quoteColumn is null ? indent : quoteColumn + "> " + indent[quoteColumn.Length..];
            var list = oe.Element(One + "List");

            // A top-level switch between bullets and numbers already starts a new list
            if (list is not null && _kind == "list" && prefix is "" or "> " && (list.Element(One + "Bullet") is not null) != _topLevelBullet)
            {
                _altMarkers = false;
            }

            // OneNote's number text can be a letter or roman numeral ("a.", "iv."): Markdown takes digits only.
            var numberText = (string?)list?.Element(One + "Number")?.Attribute("text");
            var marker = list?.Element(One + "Number") is not null
                ? (numberText is not null && NumberMarker().Match(numberText) is { Success: true } digits ? digits.Groups[1].Value : $"{number}") + (_altMarkers ? ") " : ". ")
                : list is not null ? (_altMarkers ? "* " : "- ") : null;

            var html = OeHtml(oe);
            var hasText = oe.Element(One + "T") is not null && oe != repeatedTitle;

            if (hasText && styleName == "code" && marker is null)
            {
                CodeLine(indent, Parser.ParseDocument($"<body>{html}</body>").Body!.TextContent.Replace(' ', ' '));
            }
            else if (hasText)
            {
                var heading = Regex.Match(styleName, "^h([1-6])$");

                // OneNote copies an italic style's italic onto the text, such as its own Quote, h4, and h6.
                var text = InlineMarkdown(html, (string?)oe.Attribute("style") ?? "", ignoreUniformItalic: styleItalic).Trim();
                if (marker is not null)
                {
                    ListItem(prefix, marker, EscapeLineStart(text), quote);
                }
                else if (text.Length == 0)
                {
                    Blank(indent);
                }
                else if (heading.Success)
                {
                    Paragraph(indent, $"{new string('#', int.Parse(heading.Groups[1].Value))} {EscapeClosingHashes(text)}", false);
                }
                else
                {
                    Paragraph(prefix, HorizontalRule().IsMatch(text) ? "---" : EscapeLineStart(text), quote);
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
                        Paragraph(indent, $"[Attached: {Escape((string?)child.Attribute("preferredName") ?? "file")}]", false);
                        break;
                    case "OEChildren":
                        // Only list items with a line indent their children. Indented text elsewhere would read as code.
                        Block(child, marker is null || !hasText ? indent : indent + new string(' ', marker.Length), quoteColumn);
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
            Paragraph(indent, $"[Image {reference.Index}]", false);
        }

        private void Table(XElement table, string indent)
        {
            var rows = table.Elements(One + "Row").Select((row, r) => row.Elements(One + "Cell").Select(cell =>
            {
                // The header row renders bold already
                var text = string.Join(" ", cell.Descendants(One + "OE").Where(oe => oe.Element(One + "T") is not null).Select(oe =>
                    InlineMarkdown(OeHtml(oe), (string?)oe.Attribute("style") ?? "", ignoreUniformBold: r == 0).Trim()));
                return Whitespace().Replace(text, " ").Trim().Replace("|", "\\|");
            }).ToList()).ToList();
            if (rows.Count == 0)
            {
                return;
            }

            ResetMarkersAtTopLevel(indent);
            Separate(null);
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

        /// <param name="prefix">Indent, plus "> " when quoted</param>
        private void Paragraph(string prefix, string text, bool quote)
        {
            ResetMarkersAtTopLevel(prefix);
            Separate(quote ? prefix : null);
            _lines.Add(prefix + text);
            _kind = "paragraph";
            _quotePrefix = quote ? prefix : null;
        }

        private void ListItem(string prefix, string marker, string text, bool quote)
        {
            var bullet = marker is "- " or "* ";
            var topLevel = prefix is "" or "> ";

            // A new list, a change in or out of a quote, or a top-level switch between bullets and numbers
            if (_kind != "list" || quote != _quotePrefix is not null || (topLevel && bullet != _topLevelBullet))
            {
                Separate(quote ? prefix : null);
            }

            if (topLevel)
            {
                _topLevelBullet = bullet;
            }

            _lines.Add(prefix + marker + text);
            _kind = "list";
            _quotePrefix = quote ? prefix : null;
        }

        private void CodeLine(string indent, string text)
        {
            if (_kind != "code")
            {
                if (text.Trim().Length == 0)
                {
                    return; // a blank code line outside a block
                }

                ResetMarkersAtTopLevel(indent);
                Separate(null);
                _codeIndent = indent;
                _codeStart = _lines.Count;
                _lines.Add(""); // the opening fence, set when the block closes
            }

            _lines.Add(indent + text.TrimEnd());
            _kind = "code";
        }

        /// <summary>
        /// A blank OE ends an open code block, quote, or top-level list, so back-to-back ones stay separate.
        /// </summary>
        /// <remarks>The next list after a blank uses the other markers, or Markdown would join the two lists.</remarks>
        private void Blank(string indent)
        {
            if (_kind == "list" && indent.Length == 0)
            {
                _altMarkers = !_altMarkers;
                Separate(null);
            }
            else if (_kind == "code")
            {
                Separate(null);
            }

            _quotePrefix = null;
        }

        /// <summary>
        /// Top-level content between lists ends the need for alternate markers.
        /// </summary>
        private void ResetMarkersAtTopLevel(string prefix)
        {
            if (prefix is "" or "> ")
            {
                _altMarkers = false;
            }
        }

        /// <summary>
        /// Closes an open code block and adds a blank line before the next block.
        /// </summary>
        /// <param name="nextQuotePrefix">The next block's prefix when it's quoted. Two quoted blocks in a row stay one quote.</param>
        private void Separate(string? nextQuotePrefix)
        {
            if (_kind == "code")
            {
                // A fence longer than any backtick run that starts a line inside the block
                var longest = _lines.Skip(_codeStart + 1).Select(l => l.TrimStart().TakeWhile(c => c == '`').Count()).DefaultIfEmpty(0).Max();
                var fence = _codeIndent + new string('`', Math.Max(3, longest + 1));
                _lines[_codeStart] = fence;
                _lines.Add(fence);
            }

            if (_lines.Count > 0 && _lines[^1].Length > 0)
            {
                _lines.Add(nextQuotePrefix is not null && _quotePrefix is not null && _kind != "code" ? nextQuotePrefix.TrimEnd() : "");
            }

            _kind = "";
            _quotePrefix = null;
        }

        public override string ToString()
        {
            Separate(null);
            return string.Join("\n", _lines).Trim();
        }
    }

    /// <summary>
    /// Converts the HTML inside a text element to Markdown.
    /// </summary>
    /// <param name="ignoreUniformItalic">Drop italic that covers all of the text</param>
    /// <param name="ignoreUniformBold">Drop bold that covers all of the text</param>
    private static string InlineMarkdown(string html, string paragraphStyle, bool ignoreUniformItalic = false, bool ignoreUniformBold = false)
    {
        var body = Parser.ParseDocument($"<body>{html}</body>").Body!;
        var paragraph = StyleOf(paragraphStyle);
        var ignore = (
            Bold: ignoreUniformBold && (paragraph.Bold || AllTextHas(body, (s, tag) => s.Bold || tag is "b" or "strong")),
            Italic: ignoreUniformItalic && (paragraph.Italic || AllTextHas(body, (s, tag) => s.Italic || tag is "i" or "em")));

        var text = Inline(body.ChildNodes, paragraph.Code, ignore);
        if (paragraph.Code)
        {
            return CodeSpan(text);
        }

        text = Wrap(text, paragraph.Strike ? "~~" : "");
        text = Wrap(text, paragraph.Italic && !ignore.Italic ? "*" : "");
        return Wrap(text, paragraph.Bold && !ignore.Bold ? "**" : "");
    }

    private static string Inline(IEnumerable<INode> nodes, bool inCode, (bool Bold, bool Italic) ignore)
    {
        var sb = new StringBuilder();
        var list = nodes.ToList();
        for (var i = 0; i < list.Count; i++)
        {
            var node = list[i];
            if (node is IText t)
            {
                var text = t.Data.Replace(' ', ' ');
                sb.Append(inCode ? text : Escape(Whitespace().Replace(text, " ")));
                continue;
            }

            if (node is not IElement e)
            {
                continue;
            }

            // Neighbors with the same formatting (OneNote splits runs) wrap once: **foobar**, not **foo****bar**.
            var run = new List<IElement> { e };
            while (e.LocalName is not ("a" or "br" or "img") && i + 1 < list.Count && list[i + 1] is IElement next
                && next.LocalName == e.LocalName && next.GetAttribute("style") == e.GetAttribute("style"))
            {
                run.Add(next);
                i++;
            }

            var style = StyleOf(e.GetAttribute("style") ?? "");
            var tag = e.LocalName;
            var code = inCode || style.Code || tag == "code";
            var inner = Inline(run.SelectMany(r => r.ChildNodes), code, ignore);

            if (tag == "br")
            {
                sb.Append(' ');
            }
            else if (tag == "a" && e.GetAttribute("href") is { } href && e.TextContent.Trim() == href && !href.Any(char.IsWhiteSpace))
            {
                sb.Append(Escape(href)); // OneNote links bare URLs: keep them bare
            }
            else if (tag == "a")
            {
                // Any unescaped bracket in link text would end the link early
                sb.Append($"[{UnescapedBracket().Replace(inner.Trim(), @"\$0")}]({LinkDestination(e.GetAttribute("href") ?? "")})");
            }
            else if (code && !inCode)
            {
                sb.Append(CodeSpan(inner));
            }
            else if (code)
            {
                sb.Append(inner);
            }
            else
            {
                inner = Wrap(inner, style.Strike || tag is "s" or "del" or "strike" ? "~~" : "");
                inner = Wrap(inner, (style.Italic || tag is "i" or "em") && !ignore.Italic ? "*" : "");
                sb.Append(Wrap(inner, (style.Bold || tag is "b" or "strong") && !ignore.Bold ? "**" : ""));
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// True when every non-blank text node sits inside formatting that <paramref name="has"/> matches.
    /// </summary>
    private static bool AllTextHas(IElement body, Func<(bool Bold, bool Italic, bool Strike, bool Code), string, bool> has)
    {
        var texts = body.GetDescendants().OfType<IText>().Where(t => t.Data.Trim().Length > 0).ToList();
        return texts.Count > 0 && texts.All(t =>
        {
            for (var e = t.ParentElement; e is not null && e != body; e = e.ParentElement)
            {
                if (has(StyleOf(e.GetAttribute("style") ?? ""), e.LocalName))
                {
                    return true;
                }
            }

            return false;
        });
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

    /// <summary>
    /// Inline code, with a longer fence when the code holds backticks.
    /// </summary>
    private static string CodeSpan(string code)
    {
        var core = code.Trim();
        if (core.Length == 0)
        {
            return code;
        }

        var fence = new string('`', Backticks().Matches(core).Select(m => m.Length).DefaultIfEmpty(0).Max() + 1);
        var pad = core.StartsWith('`') || core.EndsWith('`') ? " " : "";
        return Wrap(code, "").Replace(core, $"{fence}{pad}{core}{pad}{fence}");
    }

    /// <summary>
    /// A link destination Markdown reads as one: angle brackets around one with spaces or parentheses.
    /// </summary>
    private static string LinkDestination(string href)
    {
        href = href.Replace("<", "%3C").Replace(">", "%3E");
        return href.IndexOfAny([' ', '(', ')']) >= 0 ? $"<{href}>" : href;
    }

    private static (bool Bold, bool Italic, bool Strike, bool Code) StyleOf(string css)
    {
        var s = css.ToLowerInvariant().Replace(" ", "").Replace("\n", "").Replace("\r", "");
        return (s.Contains("font-weight:bold"), s.Contains("font-style:italic"), s.Contains("line-through"), Monospace().IsMatch(s));
    }

    /// <summary>
    /// Escapes characters that would read as inline Markdown.
    /// </summary>
    private static string Escape(string text) => InlineSpecial().Replace(text, @"\$0");

    /// <summary>
    /// Escapes a line start that would read as a heading, list, quote, or rule.
    /// </summary>
    private static string EscapeLineStart(string text)
    {
        if (NumberedLineStart().Match(text) is { Success: true } m)
        {
            return $"{m.Groups[1].Value}\\{text[m.Groups[1].Length..]}";
        }

        if (text.Length == 0)
        {
            return text;
        }

        var bullet = text[0] is '-' or '+' && (text.Length == 1 || char.IsWhiteSpace(text[1]));
        return text[0] is '#' or '>' || bullet || DashRule().IsMatch(text) ? "\\" + text : text;
    }

    /// <summary>
    /// The OE's text as one HTML string.
    /// </summary>
    /// <remarks>OneNote splits a paragraph into several T runs, for example around the cursor, and a T can carry its own style.</remarks>
    private static string OeHtml(XElement oe) => string.Concat(oe.Elements(One + "T").Select(t =>
        t.Attribute("style") is { } style ? $"<span style=\"{System.Net.WebUtility.HtmlEncode(style.Value)}\">{t.Value}</span>" : t.Value));

    /// <summary>
    /// Escapes a run of # at the end of heading text, which Markdown reads as a closing sequence.
    /// </summary>
    private static string EscapeClosingHashes(string text) => ClosingHashes().Replace(text, @"$1\$2");

    private static double? ToDouble(XAttribute? a) =>
        double.TryParse((string?)a, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : null;

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex Tags();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();

    [GeneratedRegex("`+")]
    private static partial Regex Backticks();

    // Spaces are already removed from the CSS
    [GeneratedRegex("font-family:[^;]*(consolas|courier|lucidaconsole|cascadia|menlo|monaco|sourcecode|firacode|firamono|jetbrainsmono|dejavusansmono|monospace)")]
    private static partial Regex Monospace();

    // \ * ` always; ] before ( [ or : (a link or link definition); _ at word edges; < before a tag; ~ in pairs; & before an entity.
    // A lone [x] stays: this output has no link definitions, so it can't become a link.
    [GeneratedRegex(@"[\\*`]|\](?=[(\[:])|(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])|<(?=[A-Za-z/!?])|~(?=~)|(?<=~)~|&(?=#?\w+;)")]
    private static partial Regex InlineSpecial();

    [GeneratedRegex(@"(?<!\\)[\[\]]")]
    private static partial Regex UnescapedBracket();

    [GeneratedRegex(@"(^|\s)(#+)\s*$")]
    private static partial Regex ClosingHashes();

    [GeneratedRegex(@"^(\d+)[.)](\s|$)")]
    private static partial Regex NumberedLineStart();

    [GeneratedRegex(@"^-[ \t]*-[ \t]*-[- \t]*$")]
    private static partial Regex DashRule();

    // Our ― rule, or an underscore line like Outlook's (escaped by now)
    [GeneratedRegex(@"^(―|\\_){3,}$")]
    private static partial Regex HorizontalRule();

    // OneNote's number text: 1. or 1)
    [GeneratedRegex(@"^(\d+)[.)]$")]
    private static partial Regex NumberMarker();
}
