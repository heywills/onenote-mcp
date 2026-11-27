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

    /// <param name="imageMarkdown">The Markdown for each image. Default: <c>[Image N]</c>.</param>
    public static (string Markdown, IReadOnlyList<ImageRef> Images) Convert(string xml, Func<ImageRef, string>? imageMarkdown = null)
    {
        var root = XElement.Parse(xml);
        var title = (string?)root.Attribute("name") ?? (string?)root.Attribute("ID") ?? "Untitled";
        // The first def wins when an index repeats
        var styles = root.Elements(One + "QuickStyleDef").GroupBy(q => (string?)q.Attribute("index") ?? "").ToDictionary(
            g => g.Key,
            g => (Name: (string?)g.First().Attribute("name") ?? "", Italic: (string?)g.First().Attribute("italic") == "true"));

        // Without a Title element, OneNote names the page after its first line, so that line repeats the name.
        var repeatedTitle = root.Element(One + "Title") is null
            ? root.Elements(One + "Outline").FirstOrDefault()?.Element(One + "OEChildren")?.Element(One + "OE")
            : null;
        if (repeatedTitle is not null && CleanText(string.Concat(repeatedTitle.Elements(One + "T").Select(t => t.Value))).Trim() != title.Trim())
        {
            repeatedTitle = null;
        }

        var writer = new Writer(styles, repeatedTitle, imageMarkdown ?? (img => $"[Image {img.Index}]"));
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

    /// <summary>Markdown for an image saved as <c>image-N.extension</c> in <paramref name="folder"/>, next to the page file.</summary>
    /// <remarks>
    /// - The folder is percent-encoded, so a page name with <c>#</c>, <c>%</c>, or spaces still links.
    /// - A type viewers can't show (<c>bin</c>) is a plain link, not an embed.
    /// </remarks>
    public static string ImageLink(int index, string folder, string extension)
    {
        var link = $"{Uri.EscapeDataString(folder)}/image-{index}.{extension}";
        return extension == "bin" ? $"[Image {index}: unsupported type]({link})" : $"![Image {index}]({link})";
    }

    /// <summary>Markdown for an image that couldn't be saved, with the reason as plain text on one line.</summary>
    public static string ImageProblem(int index, string reason) =>
        $"[Image {index}: {Escape(Whitespace().Replace(reason, " ").Trim().Replace(']', ')'))}]";

    /// <summary>
    /// Strips the inline HTML OneNote keeps inside text elements.
    /// </summary>
    public static string CleanText(string text) =>
        System.Net.WebUtility.HtmlDecode(Tags().Replace(text, "")).Replace('\u00A0', ' ');

    private sealed class Writer(Dictionary<string, (string Name, bool Italic)> styles, XElement? repeatedTitle, Func<ImageRef, string> imageMarkdown)
    {
        private readonly List<string> _lines = [];
        private string _kind = "";
        private int _codeStart;
        private string _codeIndent = "";

        /// <summary>The line prefix of the last quoted block, such as <c>"> "</c> or <c>"  > "</c>. Null when it wasn't quoted.</summary>
        private string? _quotePrefix;
        private bool _topLevelBullet;

        /// <summary>The last line is a list marker with no text.</summary>
        /// <remarks>Its content must follow on the next line: a blank line after a bare marker ends the item.</remarks>
        private bool _bareMarker;

        /// <summary>The lists open at each line prefix, and whether each is bulleted.</summary>
        private readonly Dictionary<string, bool> _openLists = [];

        /// <summary>The list indents that use <c>*</c> and <c>1)</c> markers.</summary>
        /// <remarks>Back-to-back lists of the same kind, or in side-by-side outlines, would read as one list.</remarks>
        private readonly HashSet<string> _altIndents = [];

        public List<ImageRef> Images { get; } = [];

        public void NewOutline()
        {
            var flip = _kind == "list" && !_altIndents.Contains("");
            _altIndents.Clear();
            _openLists.Clear();
            if (flip)
            {
                _altIndents.Add("");
            }

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
            // Inside an enclosing quote, "> " stays at that quote's column and the rest of the indent follows it.
            // The children of a quoted item stay in the quote, styled or not.
            if (styleName == "blockquote")
            {
                quoteColumn ??= indent;
            }

            var quote = quoteColumn is not null;
            var prefix = QuotePrefix(quoteColumn, indent);
            var marker = Marker(oe.Element(One + "List"), indent, prefix, number);
            var hasText = oe.Element(One + "T") is not null && oe != repeatedTitle;

            // A list item with no text, such as a bullet holding only an image, still gets its marker line.
            var listItem = marker is not null && oe != repeatedTitle;
            var childIndent = listItem ? indent + new string(' ', marker!.Length) : indent;
            if (listItem && !hasText)
            {
                ListItem(prefix, marker!, "", quote);
            }

            if (hasText)
            {
                Text(oe, (styleName, styleItalic), indent, prefix, marker, quote);
            }

            Children(oe, childIndent, quoteColumn);

            // A truly empty item: what follows isn't its content
            if (listItem && _lines.Count > 0 && _lines[^1] == (prefix + marker).TrimEnd())
            {
                _bareMarker = false;
            }
        }

        /// <summary>The indent, with <c>"> "</c> at the quote's column when quoted</summary>
        private static string QuotePrefix(string? quoteColumn, string indent) =>
            quoteColumn is null ? indent : quoteColumn + "> " + indent[quoteColumn.Length..];

        /// <summary>The list marker, such as <c>"- "</c> or <c>"2. "</c>. Null when the OE isn't a list item.</summary>
        private string? Marker(XElement? list, string indent, string prefix, int number)
        {
            if (list is null)
            {
                return null;
            }

            // A new item ends the deeper lists, so their alternate markers no longer apply
            _altIndents.RemoveWhere(i => i.Length > indent.Length);

            // A top-level switch between bullets and numbers already starts a new list
            if (_kind == "list" && prefix is "" or "> " && (list.Element(One + "Bullet") is not null) != _topLevelBullet)
            {
                _altIndents.Remove("");
            }

            var alt = _altIndents.Contains(indent);
            if (list.Element(One + "Number") is not { } numbered)
            {
                return alt ? "* " : "- ";
            }

            // OneNote's number text can be a letter or roman numeral ("a.", "iv."): Markdown takes digits only.
            var numberText = (string?)numbered.Attribute("text");
            var digits = numberText is not null && NumberMarker().Match(numberText) is { Success: true } match ? match.Groups[1].Value : $"{number}";
            return digits + (alt ? ") " : ". ");
        }

        private void Text(XElement oe, (string Name, bool Italic) style, string indent, string prefix, string? marker, bool quote)
        {
            var html = OeHtml(oe);
            if (style.Name == "code" && marker is null)
            {
                CodeLine(indent, Parser.ParseDocument($"<body>{html}</body>").Body!.TextContent.Replace('\u00A0', ' '));
                return;
            }

            // OneNote copies an italic style's italic onto the text, such as its own Quote, h4, and h6.
            var text = InlineMarkdown(html, (string?)oe.Attribute("style") ?? "", ignoreUniformItalic: style.Italic).Trim();
            var heading = Regex.Match(style.Name, "^h([1-6])$");
            if (marker is not null)
            {
                ListItem(prefix, marker, EscapeLineStart(text), quote);
            }
            else if (text.Length == 0)
            {
                Blank(indent, prefix, quote);
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

        /// <summary>Writes an OE's images, tables, attached files, and nested OEs.</summary>
        /// <remarks>
        /// - Only list items indent their content. Indented text elsewhere would read as code.
        /// </remarks>
        private void Children(XElement oe, string childIndent, string? quoteColumn)
        {
            var quote = quoteColumn is not null;
            var childPrefix = QuotePrefix(quoteColumn, childIndent);
            foreach (var child in oe.Elements())
            {
                switch (child.Name.LocalName)
                {
                    case "Image":
                        Image(child, childPrefix, quote);
                        break;
                    case "Table":
                        Table(child, childIndent);
                        break;
                    case "InsertedFile":
                        Paragraph(childPrefix, $"[Attached: {Escape((string?)child.Attribute("preferredName") ?? "file")}]", quote);
                        break;
                    case "OEChildren":
                        Block(child, childIndent, quoteColumn);
                        break;
                }
            }
        }

        /// <param name="prefix">Indent, plus "> " when quoted</param>
        public void Image(XElement img, string prefix, bool quote = false)
        {
            var id = (string?)img.Element(One + "CallbackID")?.Attribute("callbackID") ?? (string?)img.Attribute("callbackID");
            if (id is null)
            {
                return;
            }

            // An image shown again reuses its number
            var reference = Images.FirstOrDefault(i => i.CallbackId == id);
            if (reference is null)
            {
                var size = img.Element(One + "Size");
                reference = new ImageRef(id, Images.Count + 1, ToDouble(size?.Attribute("width")), ToDouble(size?.Attribute("height")));
                Images.Add(reference);
            }

            Paragraph(prefix, imageMarkdown(reference), quote);
        }

        private void Table(XElement table, string indent)
        {
            var rows = table.Elements(One + "Row").Select((row, r) => row.Elements(One + "Cell").Select(cell =>
            {
                // The header row renders bold already
                // A code-style line (from <pre>) is inline code: a cell holds one line.
                var text = string.Join(" ", cell.Descendants(One + "OE").Where(oe => oe.Element(One + "T") is not null).Select(oe =>
                    styles.GetValueOrDefault((string?)oe.Attribute("quickStyleIndex") ?? "").Name == "code"
                        ? CodeSpan(Parser.ParseDocument($"<body>{OeHtml(oe)}</body>").Body!.TextContent.Replace('\u00A0', ' '))
                        : InlineMarkdown(OeHtml(oe), (string?)oe.Attribute("style") ?? "", ignoreUniformBold: r == 0).Trim()));
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

            else if (!bullet && marker.TrimEnd()[..^1] != "1" && !_bareMarker
                && (!_openLists.TryGetValue(prefix, out var openBullet) || openBullet))
            {
                // A new numbered list that starts above 1 can't interrupt the line above it, so it needs a blank line first
                _lines.Add(prefix.TrimEnd());
            }

            if (topLevel)
            {
                _topLevelBullet = bullet;
            }

            // Deeper lists end when an item at this indent starts
            foreach (var deeper in _openLists.Keys.Where(k => k.Length > prefix.Length).ToList())
            {
                _openLists.Remove(deeper);
            }

            _openLists[prefix] = bullet;
            _lines.Add((prefix + marker + text).TrimEnd());
            _kind = "list";
            _quotePrefix = quote ? prefix : null;
            _bareMarker = text.Length == 0;
        }

        private void CodeLine(string indent, string text)
        {
            // A block at another indent is another block, such as one in a list item and one after the list
            if (_kind == "code" && indent != _codeIndent)
            {
                Separate(null);
            }

            if (_kind != "code")
            {
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
        /// A blank OE ends an open code block, a list, or a quote, so back-to-back ones stay separate.
        /// </summary>
        /// <remarks>
        /// - The next list at that indent uses the other markers, or Markdown would join the two lists.
        /// - A blank OE in the quote style keeps the quote going.
        /// </remarks>
        /// <param name="prefix">The blank OE's indent, plus "> " when quoted: the list it ends is open at that prefix</param>
        private void Blank(string indent, string prefix, bool quote)
        {
            // A list is open here even when its last item ended in a child paragraph, image, code, or table
            if (_openLists.ContainsKey(prefix))
            {
                if (!_altIndents.Remove(indent))
                {
                    _altIndents.Add(indent);
                }

                foreach (var closed in _openLists.Keys.Where(k => k.Length >= prefix.Length).ToList())
                {
                    _openLists.Remove(closed);
                }

                if (_kind == "list")
                {
                    _kind = "listBreak"; // the next list item separates
                }
            }
            else if (_kind == "code")
            {
                Separate(null);
            }

            if (!quote)
            {
                _quotePrefix = null;
            }
        }

        /// <summary>
        /// Top-level content between lists ends the need for alternate markers.
        /// </summary>
        private void ResetMarkersAtTopLevel(string prefix)
        {
            if (prefix is "" or "> ")
            {
                _altIndents.Clear();
                _openLists.Clear();
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
                // An empty block is one blank code line in OneNote
                if (_lines.Skip(_codeStart + 1).All(l => l.Trim().Length == 0))
                {
                    _lines.RemoveRange(_codeStart + 1, _lines.Count - _codeStart - 1);
                }

                var longest = _lines.Skip(_codeStart + 1).Select(l => l.TrimStart().TakeWhile(c => c == '`').Count()).DefaultIfEmpty(0).Max();
                var fence = _codeIndent + new string('`', Math.Max(3, longest + 1));
                _lines[_codeStart] = fence;
                _lines.Add(fence);
            }

            if (_lines.Count > 0 && _lines[^1].Length > 0 && !(_bareMarker && _kind == "list"))
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
            Italic: ignoreUniformItalic && (paragraph.Italic || AllTextHas(body, (s, tag) => s.Italic || tag is "i" or "em")),
            Strike: false);

        // Formatting the paragraph already wraps doesn't wrap again inside it
        var text = Inline(body.ChildNodes, paragraph.Code, (ignore.Bold || paragraph.Bold, ignore.Italic || paragraph.Italic, paragraph.Strike));
        if (paragraph.Code)
        {
            return HasLink(body) ? CodeWithLinks(body.ChildNodes) : CodeSpan(text);
        }

        text = Wrap(text, paragraph.Strike ? "~~" : "");
        text = Wrap(text, paragraph.Italic && !ignore.Italic ? "*" : "");
        return Wrap(text, paragraph.Bold && !ignore.Bold ? "**" : "");
    }

    /// <param name="ignore">Formatting to skip: dropped, or already wrapped by an enclosing element</param>
    /// <param name="inLink">Inside link text, where any bracket outside code would end the link</param>
    /// <param name="outerBefore">The character before these nodes, for emphasis flanking</param>
    /// <param name="outerAfter">The character after these nodes, for emphasis flanking</param>
    private static string Inline(IEnumerable<INode> nodes, bool inCode, (bool Bold, bool Italic, bool Strike) ignore, bool inLink = false, char outerBefore = ' ', char outerAfter = ' ')
    {
        var sb = new StringBuilder();
        var list = nodes.ToList();
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i] is IText t)
            {
                sb.Append(TextMarkdown(t.Data, inCode, inLink));
            }
            else if (list[i] is IElement)
            {
                var run = TakeRun(list, ref i, inCode);

                // Markers right after another run's closing marker merge with it ("*foo***.bar**"), so use HTML there
                var afterMarker = sb.Length > 0 && sb[^1] is '*' or '~' && (sb.Length < 2 || sb[^2] != '\\');

                // The next character, past any empty elements
                var neighbors = (Before: sb.Length > 0 ? sb[^1] : outerBefore, After: NextChar(list.Skip(i + 1), outerAfter), afterMarker);
                sb.Append(RunMarkdown(run, inCode, ignore, inLink, neighbors));
            }
        }

        return sb.ToString();
    }

    private static string TextMarkdown(string data, bool inCode, bool inLink)
    {
        var text = data.Replace('\u00A0', ' ');
        if (inCode)
        {
            return text;
        }

        text = Whitespace().Replace(text, " ");
        return inLink ? LinkTextSpecial().Replace(text, @"\$0") : Escape(text);
    }

    /// <summary>
    /// The element at <paramref name="i"/> and its neighbors with the same formatting. Moves <paramref name="i"/> to the last one.
    /// </summary>
    /// <remarks>
    /// - Neighbors with the same formatting wrap once: **foobar**, not **foo****bar**.
    /// - OneNote splits runs where only the color or font changes, so compare formatting, not the whole style.
    /// - Code runs always merge, or `a``b` would read as one span holding a double backtick.
    /// </remarks>
    private static List<IElement> TakeRun(List<INode> list, ref int i, bool inCode)
    {
        var e = (IElement)list[i];
        var run = new List<IElement> { e };
        while (e.LocalName is not ("a" or "br" or "img") && i + 1 < list.Count && list[i + 1] is IElement next
            && next.LocalName is not ("a" or "br" or "img")
            && (Formatting(next) == Formatting(e) || Formatting(next).Code && Formatting(e).Code && !inCode))
        {
            run.Add(next);
            i++;
        }

        return run;
    }

    /// <param name="neighbors">The characters around the run, and whether it follows another run's closing marker</param>
    private static string RunMarkdown(List<IElement> run, bool inCode, (bool Bold, bool Italic, bool Strike) ignore, bool inLink, (char Before, char After, bool AfterMarker) neighbors)
    {
        var e = run[0];
        var style = StyleOf(e.GetAttribute("style") ?? "");
        var tag = e.LocalName;
        var code = inCode || style.Code || tag == "code";
        var strike = (style.Strike || tag is "s" or "del" or "strike") && !ignore.Strike;
        var italic = (style.Italic || tag is "i" or "em") && !ignore.Italic;
        var bold = (style.Bold || tag is "b" or "strong") && !ignore.Bold;
        var inner = Inline(run.SelectMany(r => r.ChildNodes), code, (ignore.Bold || bold, ignore.Italic || italic, ignore.Strike || strike), inLink || tag == "a", neighbors.Before, neighbors.After);

        if (tag == "br")
        {
            return " ";
        }

        if (tag == "a")
        {
            return LinkMarkdown(e, inner, code, inCode);
        }

        if (code && !inCode)
        {
            // Merged code runs keep formatting they all share: **`x`**
            var formats = run.Select(Formatting).ToList();
            var codeMarkdown = run.Any(HasLink) ? CodeWithLinks(run.SelectMany(r => r.ChildNodes)) : CodeSpan(inner);
            return Emphasize(codeMarkdown, formats.All(f => f.Strike) && !ignore.Strike, formats.All(f => f.Italic) && !ignore.Italic,
                formats.All(f => f.Bold) && !ignore.Bold, neighbors.Before, neighbors.After, neighbors.AfterMarker);
        }

        return code ? inner : Emphasize(inner, strike, italic, bold, neighbors.Before, neighbors.After, neighbors.AfterMarker);
    }

    private static string LinkMarkdown(IElement a, string inner, bool code, bool inCode)
    {
        var href = a.GetAttribute("href");
        if (inCode)
        {
            return inner; // a code span can't hold a link: keep the text
        }

        if (href is not null && a.TextContent.Trim() == href && AutolinkUrl().IsMatch(href))
        {
            return $"<{href}>"; // OneNote links bare URLs. An autolink needs no escaping.
        }

        if (string.IsNullOrEmpty(href))
        {
            return inner; // an anchor, not a link
        }

        // A monospace link: [`text`](url)
        var text = code ? CodeSpan(inner.Trim()) : inner.Trim();
        return $"[{text}]({LinkDestination(href)})";
    }

    private static bool HasLink(IElement e) => e.QuerySelector("a[href]") is not null;

    /// <summary>
    /// Monospace text holding links, as code spans around <c>[`text`](url)</c> links: a code span can't hold a link.
    /// </summary>
    private static string CodeWithLinks(IEnumerable<INode> nodes)
    {
        var sb = new StringBuilder();
        var code = new StringBuilder();

        void Flush()
        {
            sb.Append(CodeSpan(code.ToString()));
            code.Clear();
        }

        void Walk(IEnumerable<INode> children)
        {
            foreach (var node in children)
            {
                if (node is IElement { LocalName: "a" } a && !string.IsNullOrEmpty(a.GetAttribute("href")))
                {
                    Flush();
                    sb.Append($"[{CodeSpan(a.TextContent.Replace('\u00A0', ' ').Trim())}]({LinkDestination(a.GetAttribute("href")!)})");
                }
                else if (node is IElement element && HasLink(element))
                {
                    Walk(element.ChildNodes);
                }
                else
                {
                    code.Append(node is IElement { LocalName: "br" } ? " " : node.TextContent.Replace('\u00A0', ' '));
                }
            }
        }

        Walk(nodes);
        Flush();
        return sb.ToString();
    }

    /// <summary>
    /// Wraps text in strikethrough, italic, and bold.
    /// </summary>
    /// <remarks>
    /// - "func**(x)**" shows the asterisks: Markdown emphasis can't open between a letter and punctuation, so those cases use HTML.
    /// - So does text right after another run's closing marker, where the markers would merge.
    /// </remarks>
    private static string Emphasize(string text, bool strike, bool italic, bool bold, char before, char after, bool afterMarker = false)
    {
        if (afterMarker && (strike || italic || bold) || !MarkersCanWrap(text, before, after))
        {
            text = WrapHtml(text, strike ? "del" : "");
            text = WrapHtml(text, italic ? "em" : "");
            return WrapHtml(text, bold ? "strong" : "");
        }

        text = Wrap(text, strike ? "~~" : "");
        text = Wrap(text, italic ? "*" : "");
        return Wrap(text, bold ? "**" : "");
    }

    /// <summary>
    /// The Markdown formatting an element carries: bold, italic, strikethrough, code.
    /// </summary>
    private static (bool Bold, bool Italic, bool Strike, bool Code) Formatting(IElement e)
    {
        var s = StyleOf(e.GetAttribute("style") ?? "");
        return (s.Bold || e.LocalName is "b" or "strong", s.Italic || e.LocalName is "i" or "em",
            s.Strike || e.LocalName is "s" or "del" or "strike", s.Code || e.LocalName == "code");
    }

    /// <summary>
    /// Whether Markdown emphasis markers around <paramref name="text"/> would open and close, given its neighbors.
    /// </summary>
    /// <remarks>
    /// - A marker next to punctuation counts only with whitespace or punctuation on its other side (CommonMark flanking rules).
    /// - Wrap keeps edge spaces outside, so a space at an edge makes that edge's neighbor a space.
    /// </remarks>
    private static bool MarkersCanWrap(string text, char before, char after)
    {
        var core = text.Trim();
        if (core.Length == 0)
        {
            return true;
        }

        if (text[0] == ' ')
        {
            before = ' ';
        }

        if (text[^1] == ' ')
        {
            after = ' ';
        }

        var opens = !char.IsPunctuation(core[0]) && !char.IsSymbol(core[0]) || char.IsWhiteSpace(before) || char.IsPunctuation(before) || char.IsSymbol(before);
        var closes = !char.IsPunctuation(core[^1]) && !char.IsSymbol(core[^1]) || char.IsWhiteSpace(after) || char.IsPunctuation(after) || char.IsSymbol(after);
        return opens && closes;
    }

    /// <summary>
    /// Wraps text in an HTML tag, keeping edge spaces outside.
    /// </summary>
    private static string WrapHtml(string text, string tag)
    {
        var core = text.Trim();
        if (tag.Length == 0 || core.Length == 0)
        {
            return text;
        }

        return $"{text[..(text.Length - text.TrimStart().Length)]}<{tag}>{core}</{tag}>{text[text.TrimEnd().Length..]}";
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
    /// <remarks>
    /// - Backslashes are doubled: Markdown reads one before punctuation as an escape, which breaks \\server\share links.
    /// - OneNote page links come back in the "Copy link to page" form.
    /// </remarks>
    private static string LinkDestination(string href)
    {
        href = OneNoteLink(href).Replace("\\", "\\\\").Replace("<", "%3C").Replace(">", "%3E");
        return href.IndexOfAny([' ', '(', ')']) >= 0 ? $"<{href}>" : href;
    }

    /// <summary>
    /// A OneNote page link in the "Copy link to page" form, <c>onenote:https://…/Section.one#Page&amp;section-id={…}&amp;page-id={…}&amp;end</c>.
    /// </summary>
    /// <remarks>OneNote stores links within a notebook as <c>onenote:#Page&amp;…&amp;base-path=https://…/Section.one</c>, which onenote_get_page_by_link can't open.</remarks>
    private static string OneNoteLink(string href)
    {
        var m = RelativeOneNoteLink().Match(href);
        return m.Success ? $"onenote:{m.Groups[2].Value}#{Uri.UnescapeDataString(m.Groups[1].Value)}" : href;
    }

    /// <summary>
    /// The first character of the next nodes, past empty elements. A line break counts as a space.
    /// </summary>
    private static char NextChar(IEnumerable<INode> nodes, char fallback)
    {
        foreach (var node in nodes)
        {
            if (node is IElement { LocalName: "br" })
            {
                return ' ';
            }

            if (node.TextContent.Length > 0)
            {
                return node.TextContent[0];
            }
        }

        return fallback;
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

    /// <summary>A monospace font-family. Spaces are already removed from the CSS.</summary>
    [GeneratedRegex("font-family:[^;]*(consolas|courier|lucidaconsole|cascadia|menlo|monaco|sourcecode|firacode|firamono|jetbrainsmono|dejavusansmono|monospace)")]
    private static partial Regex Monospace();

    /// <summary>Characters that would read as inline Markdown.</summary>
    /// <remarks>
    /// - <c>\ * `</c> always; <c>]</c> before <c>( [ :</c> (a link or link definition); <c>_</c> at word edges; <c>&lt;</c> before a tag; <c>~</c> in pairs; <c>&amp;</c> before an entity
    /// - A lone <c>[x]</c> stays: this output has no link definitions, so it can't become a link.
    /// </remarks>
    [GeneratedRegex(@"[\\*`]|\](?=[(\[:])|(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])|<(?=[A-Za-z/!?])|~(?=~)|(?<=~)~|&(?=#?\w+;)")]
    private static partial Regex InlineSpecial();

    /// <summary><see cref="InlineSpecial"/> plus every bracket, which would end link text early.</summary>
    [GeneratedRegex(@"[\\*`\[\]]|(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])|<(?=[A-Za-z/!?])|~(?=~)|(?<=~)~|&(?=#?\w+;)")]
    private static partial Regex LinkTextSpecial();

    [GeneratedRegex(@"(^|\s)(#+)\s*$")]
    private static partial Regex ClosingHashes();

    [GeneratedRegex(@"^(\d+)[.)](\s|$)")]
    private static partial Regex NumberedLineStart();

    [GeneratedRegex(@"^-[ \t]*-[ \t]*-[- \t]*$")]
    private static partial Regex DashRule();

    /// <summary>Our ― rule, or an underscore line like Outlook's (escaped by now).</summary>
    [GeneratedRegex(@"^(―|\\_){3,}$")]
    private static partial Regex HorizontalRule();

    /// <summary>A URL a CommonMark autolink can hold: a scheme, and no spaces, angle brackets, or backslashes.</summary>
    [GeneratedRegex(@"^[A-Za-z][A-Za-z0-9+.-]{1,31}:[^\s<>\\]*$")]
    private static partial Regex AutolinkUrl();

    /// <summary>A OneNote link stored relative to its section: the fragment, then the base path.</summary>
    [GeneratedRegex(@"^onenote:#(.*?)&base-path=(.+)$")]
    private static partial Regex RelativeOneNoteLink();

    /// <summary>OneNote's number text: <c>1.</c> or <c>1)</c></summary>
    [GeneratedRegex(@"^(\d+)[.)]$")]
    private static partial Regex NumberMarker();
}
