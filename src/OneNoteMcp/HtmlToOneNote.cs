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
/// - <c>p</c>, <c>div</c>: one OE per paragraph. From Markdown, a <c>br</c> starts a new one.
/// - <c>h1</c>-<c>h6</c>: OE with a heading QuickStyleDef. From Markdown, a blank OE before it.
/// - A top-level <c>div class="outline"</c> with <c>left</c> and <c>top</c>: an Outline of its own at that place, such as a margin note
/// - <c>ul</c>, <c>ol</c>, <c>li</c>: OE with <c>one:List</c>, nested lists in OEChildren
/// - <c>table</c>: <c>one:Table</c>
/// - <c>pre</c>: one OE per line, in OneNote's Code style
/// - <c>blockquote</c>: paragraphs and lists in OneNote's Quote style. Headings, tables, and code inside a quote lose the quote.
/// </remarks>
public static partial class HtmlToOneNote
{
    /// <summary>OneNote's built-in styles: (font, size, color, italic)</summary>
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

    /// <summary>The most image data one page embeds</summary>
    public const long MaxPageImageBytes = 50L * 1024 * 1024;

    /// <summary>Elements whose text a page never shows, such as those from a saved web page's head</summary>
    private static readonly HashSet<string> NonContent =
        ["script", "style", "title", "template", "meta", "link", "noscript", "noframes", "iframe"];

    /// <summary>
    /// Appends <paramref name="html"/> to a page as an Outline, and adds the styles it needs.
    /// </summary>
    /// <param name="fromMarkdown">The HTML came from Markdown, whose table code keeps <c>\|</c> for a pipe</param>
    /// <param name="loadImage">Returns an image's bytes from its <c>src</c>, up to a size, or null to keep it a link</param>
    /// <returns>How many images were embedded, and how many stayed links</returns>
    public static (int Embedded, int Linked) AddToPage(XElement page, string html, bool fromMarkdown = false, Func<string, long, byte[]?>? loadImage = null)
    {
        // Parsed as body content even for a whole document: the parser ignores stray html, head, and body tags
        var body = new HtmlParser().ParseDocument($"<body>{html}</body>").Body!;
        var defs = page.Elements(One + "QuickStyleDef").ToList();

        // Reuse the page's styles by name. New ones get the next free indexes.
        var indexes = defs.Where(d => d.Attribute("name") is not null && int.TryParse((string?)d.Attribute("index"), out _))
            .GroupBy(d => (string)d.Attribute("name")!)
            .ToDictionary(g => g.Key, g => int.Parse((string)g.First().Attribute("index")!));
        var builder = new Builder(indexes, defs.Select(d => int.TryParse((string?)d.Attribute("index"), out var i) ? i : 0).DefaultIfEmpty(-1).Max() + 1, fromMarkdown, loadImage);
        var tagDefs = page.Elements(One + "TagDef").ToList();
        builder.UseTags(PageTags.Definitions(page));

        // Placed outlines, as onenote_get_page's HTML writes margin notes, become outlines of their own
        var placed = body.Children.Where(IsPlacedOutline).ToList();
        placed.ForEach(e => e.Remove());
        var outlines = new List<XElement> { NewOutline(builder, body.ChildNodes, null, 0) };
        // A div without a z-index stacks above every one with
        var topZ = placed.Select(div => ZIndex(div.GetAttribute("style"))).Max() ?? 0;
        outlines.AddRange(placed.Select((div, i) => NewOutline(builder, div.ChildNodes, div.GetAttribute("style"), topZ + i + 1)));

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

        AddDefinitions(page, tagDefs.LastOrDefault(), builder.NewTags, defs.LastOrDefault(), newDefs);

        // OneNote rejects an Outline with nothing in it
        page.Add(outlines.Where(o => o.Element(One + "OEChildren")!.HasElements));
        return (builder.Embedded, builder.Linked);
    }

    /// <summary>A <c>z-index</c> up to a million, so the ones numbered after it can't overflow.</summary>
    private static int? ZIndex(string? style) => int.TryParse(Css.Value(style, "z-index"), out var z) && z is >= 0 and <= 1_000_000 ? z : null;

    /// <summary>A top-level <c>div class="outline"</c> with a <c>left</c> and <c>top</c> in points. Other divs, even with that class, are ordinary blocks.</summary>
    private static bool IsPlacedOutline(IElement e) =>
        e.LocalName == "div" && e.ClassList.Contains("outline")
        && Coordinate(Css.Value(e.GetAttribute("style"), "left")) is not null && Coordinate(Css.Value(e.GetAttribute("style"), "top")) is not null;

    /// <summary>An Outline holding <paramref name="nodes"/>, at the place <paramref name="style"/> gives, if any.</summary>
    /// <param name="order">The outline's stacking order when the style has no <c>z-index</c></param>
    /// <remarks>
    /// - OneNote needs a Size with a height, and a Position with z. It sizes the height to the content itself.
    /// - A width means a fixed width, as onenote_get_page writes only those. Without one, the outline grows to fit.
    /// </remarks>
    private static XElement NewOutline(Builder builder, IEnumerable<INode> nodes, string? style, int order)
    {
        var children = new XElement(One + "OEChildren");
        builder.Blocks(nodes, children);
        var outline = new XElement(One + "Outline");
        if (Coordinate(Css.Value(style, "left")) is { } x && Coordinate(Css.Value(style, "top")) is { } y)
        {
            var z = ZIndex(style) ?? order;
            outline.Add(new XElement(One + "Position", new XAttribute("x", x), new XAttribute("y", y), new XAttribute("z", z)));
            if (Points(Css.Value(style, "width")) is { } width)
            {
                outline.Add(new XElement(One + "Size", new XAttribute("width", width), new XAttribute("height", "20"), new XAttribute("isSetByUser", "true")));
            }
        }

        outline.Add(children);
        return outline;
    }

    /// <summary>Adds new TagDefs and QuickStyleDefs in schema order: TagDefs, then QuickStyleDefs, then the rest.</summary>
    /// <param name="lastTag">The page's last TagDef, or null</param>
    /// <param name="lastStyle">The page's last QuickStyleDef, or null</param>
    private static void AddDefinitions(XElement page, XElement? lastTag, List<XElement> newTags, XElement? lastStyle, List<XElement> newStyles)
    {
        if ((lastStyle ?? lastTag) is { } beforeStyles)
        {
            beforeStyles.AddAfterSelf(newStyles);
        }
        else
        {
            page.AddFirst(newStyles);
        }

        // After the page's TagDefs, which are still ahead of every QuickStyleDef
        if (lastTag is not null)
        {
            lastTag.AddAfterSelf(newTags);
        }
        else
        {
            page.AddFirst(newTags);
        }
    }

    /// <summary>The document's <c>&lt;title&gt;</c>, or null when it has none.</summary>
    public static string? Title(string html) =>
        new HtmlParser().ParseDocument(html).Title?.Trim() is { Length: > 0 } title ? title : null;

    private static bool IsInline(INode n) => n is IText || (n is IElement e && Inline.Contains(e.LocalName));

    /// <summary>
    /// Renders inline nodes as the HTML OneNote accepts inside <c>one:T</c>.
    /// </summary>
    /// <param name="inCode">Keep every space, as <c>&amp;nbsp;</c></param>
    /// <param name="inLink">Inside a link, where an image can only be its alt text</param>
    /// <param name="inTable">In a table cell, where code keeps the <c>\|</c> that Markdown needs for a pipe</param>
    /// <param name="imageLinked">Called for each image written as a link or alt text, to count them</param>
    private static string InlineHtml(IEnumerable<INode> nodes, bool inCode = false, bool inLink = false, bool inTable = false, Action? imageLinked = null)
    {
        var sb = new StringBuilder();
        foreach (var n in nodes)
        {
            if (n is IText t)
            {
                // Inline code is one line: a newline in it would split the OneNote text
                sb.Append(inCode
                    ? CodeHtml((inTable ? t.Data.Replace(@"\|", "|") : t.Data).Replace("\r", "").Replace('\n', ' '))
                    : WebUtility.HtmlEncode(Whitespace().Replace(t.Data, " ")));
                continue;
            }

            if (n is not IElement e || NonContent.Contains(e.LocalName))
            {
                continue;
            }

            var inner = InlineHtml(e.ChildNodes, inCode || e.LocalName == "code", inLink || e.LocalName == "a", inTable, imageLinked);
            if (e.LocalName == "img")
            {
                imageLinked?.Invoke();
            }

            var alt = e.GetAttribute("alt") is { Length: > 0 } a ? a : e.GetAttribute("src") ?? "image";
            sb.Append(e.LocalName switch
            {
                "b" or "strong" => $"<b>{inner}</b>",
                "i" or "em" => $"<i>{inner}</i>",
                "u" => $"<u>{inner}</u>",
                "s" or "strike" or "del" => $"<span style='text-decoration:line-through'>{inner}</span>",
                "sup" or "sub" => $"<{e.LocalName}>{inner}</{e.LocalName}>",

                // OneNote drops a bare <br> in text, but keeps one followed by a newline
                "br" => "<br>\n",
                "code" => $"<span style='font-family:Consolas'>{inner}</span>",
                "a" when SafeHtml.IsSafeHref(e.GetAttribute("href") ?? "") => $"<a href=\"{WebUtility.HtmlEncode(Href(e.GetAttribute("href")))}\">{inner}</a>",

                // Colors, highlights, fonts, and sizes, such as onenote_get_page's HTML keeps
                "span" when Css.Inline(e.GetAttribute("style")) is { Length: > 0 } css => $"<span style=\"{WebUtility.HtmlEncode(css)}\">{inner}</span>",

                // OneNote can't load an image from a URL here: keep it as a link, or its alt text inside a link, with no src, or for a script link.
                "img" when inLink || e.GetAttribute("src") is not { Length: > 0 } src || !SafeHtml.IsSafeHref(src) => WebUtility.HtmlEncode(alt),
                "img" => $"<a href=\"{WebUtility.HtmlEncode(Href(e.GetAttribute("src")))}\">{WebUtility.HtmlEncode(alt)}</a>",
                _ => inner,
            });
        }

        return sb.ToString();
    }

    /// <summary>
    /// Code text as OneNote HTML: every space and tab kept as an entity, so trimming can't drop indentation.
    /// </summary>
    private static string CodeHtml(string text) => WebUtility.HtmlEncode(text).Replace(" ", "&nbsp;").Replace("\t", "&#9;");

    /// <summary>
    /// The link target, with backslashes back.
    /// </summary>
    /// <remarks>
    /// - Markdig writes a backslash in a URL as <c>%5C</c>, which breaks <c>\\server\share</c> links in OneNote.
    /// - Only file-share and file: links: elsewhere %5C is part of the URL.
    /// </remarks>
    private static string Href(string? href)
    {
        href ??= "";
        return href.StartsWith("%5C%5C", StringComparison.OrdinalIgnoreCase) || href.StartsWith(@"\\") || href.StartsWith("file:", StringComparison.OrdinalIgnoreCase)
            ? href.Replace("%5C", "\\", StringComparison.OrdinalIgnoreCase)
            : href;
    }

    /// <param name="indexes">The page's QuickStyleDef indexes by name</param>
    /// <param name="nextIndex">The first free QuickStyleDef index</param>
    /// <param name="fromMarkdown">Turn <c>\|</c> in table code back into a pipe</param>
    /// <param name="loadImage">Returns an image's bytes from its <c>src</c>, up to a size, or null to keep it a link</param>
    private sealed class Builder(Dictionary<string, int> indexes, int nextIndex, bool fromMarkdown, Func<string, long, byte[]?>? loadImage)
    {
        public int Embedded { get; private set; }

        public int Linked { get; private set; }

        private void CountLink() => Linked++;

        /// <remarks>
        /// - From HTML, <c>br</c> is a line break inside the paragraph. From Markdown, a hard break splits the paragraph.
        /// </remarks>
        private bool IsInlineNode(INode n) => IsInline(n) || (!fromMarkdown && n is IElement { LocalName: "br" });

        private bool _inQuote;
        private bool _inTable;

        /// <summary>Per parent OEChildren: whether its next numbered list uses <c>##)</c>.</summary>
        /// <remarks>
        /// - OneNote continues numbering across sibling lists with the same format, even with paragraphs between them. Alternating formats restarts each list at 1.
        /// - Lists under different parents restart on their own.
        /// </remarks>
        private readonly Dictionary<XElement, bool> _altNumberFormat = [];

        /// <summary>Styles this content needs that the page doesn't have yet</summary>
        public List<(string Name, int Index)> NewStyles { get; } = [];

        /// <summary>Tag definition indexes, by tag with <c>Done</c> false</summary>
        private readonly Dictionary<Tag, int> _tagIndexes = [];
        private int _nextTagIndex;

        /// <summary>TagDefs this content needs that the page doesn't have yet</summary>
        public List<XElement> NewTags { get; } = [];

        /// <summary>Reuses the page's tag definitions, by index.</summary>
        public void UseTags(Dictionary<string, Tag> definitions)
        {
            foreach (var (index, tag) in definitions)
            {
                if (int.TryParse(index, out var i))
                {
                    _tagIndexes.TryAdd(tag, i);
                    _nextTagIndex = Math.Max(_nextTagIndex, i + 1);
                }
            }
        }

        private int TagIndex(Tag tag)
        {
            var key = tag with { Done = false };
            if (!_tagIndexes.TryGetValue(key, out var index))
            {
                index = _tagIndexes[key] = _nextTagIndex++;
                NewTags.Add(PageTags.Definition(key, index));
            }

            return index;
        }

        /// <summary>Removes the tag checkboxes and labels at the start of a block, and returns their tags.</summary>
        /// <remarks>
        /// - onenote_get_page's HTML marks each with <c>data-tag</c>. A Markdown task item's checkbox has none, and is a To Do.
        /// - Other HTML's checkboxes without <c>data-tag</c> aren't tags.
        /// </remarks>
        private List<Tag> TakeTags(IElement block)
        {
            var tags = new List<Tag>();
            foreach (var node in block.ChildNodes.ToList())
            {
                if (node is IText text && string.IsNullOrWhiteSpace(text.Data))
                {
                    continue;
                }

                if (node is not IElement e || TagOf(e) is not { } tag)
                {
                    break;
                }

                tags.Add(tag);
                e.Remove();
            }

            return tags;
        }

        private Tag? TagOf(IElement e) => e.LocalName switch
        {
            "input" when string.Equals(e.GetAttribute("type"), "checkbox", StringComparison.OrdinalIgnoreCase) && e.HasAttribute("data-tag") =>
                DataTag(e, e.HasAttribute("checked"), Tag.ToDo(false)),
            "input" when string.Equals(e.GetAttribute("type"), "checkbox", StringComparison.OrdinalIgnoreCase) && fromMarkdown =>
                Tag.ToDo(e.HasAttribute("checked")),
            "span" when e.HasAttribute("data-tag") => DataTag(e, e.GetAttribute("data-tag-done") == "true", new Tag("Tag", 0, 0, false)),
            _ => null,
        };

        /// <param name="fallback">The type and symbol when the attributes are missing</param>
        private static Tag DataTag(IElement e, bool done, Tag fallback) => new(
            e.GetAttribute("data-tag") is { Length: > 0 } name ? name : fallback.Name,
            int.TryParse(e.GetAttribute("data-tag-type"), out var type) ? type : fallback.Type,
            int.TryParse(e.GetAttribute("data-tag-symbol"), out var symbol) ? symbol : fallback.Symbol,
            done);

        private void AddTags(XElement? oe, List<Tag> tags)
        {
            if (oe is not null && tags.Count > 0)
            {
                PageTags.Add(oe, tags, TagIndex);
            }
        }

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
        /// <param name="content">An image, in place of the text</param>
        private static XElement Oe(XElement parent, string innerHtml, int? style = null, string? listKind = null, (int Value, string Format) number = default, XElement? content = null)
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

            oe.Add(content ?? new XElement(One + "T", new XCData(innerHtml.Trim())));
            parent.Add(oe);
            return oe;
        }

        private long _embeddedBytes;

        /// <summary>The OneNote image for an <c>img</c> whose file loads. Null for anything else, which stays a link.</summary>
        /// <remarks>
        /// - Images past <see cref="MaxPageImageBytes"/> in total stay links, so one page update stays a size OneNote accepts.
        /// </remarks>
        private XElement? EmbeddedImage(INode n)
        {
            if (n is not IElement { LocalName: "img" } img)
            {
                return null;
            }

            // A cautious bound: with less than MaxEmbedKb left, skip files larger than what's left, even ones that might shrink to fit
            var remaining = MaxPageImageBytes - _embeddedBytes;
            var readLimit = remaining >= Images.MaxEmbedKb * 1024L ? long.MaxValue : remaining;
            if (remaining <= 0 || loadImage?.Invoke(img.GetAttribute("src") ?? "", readLimit) is not { } data
                || Images.ForOneNote(data) is not { } image || image.Data.Length > remaining)
            {
                return null;
            }

            _embeddedBytes += image.Data.Length;
            Embedded++;
            // Schema order: Size before Data
            return new XElement(One + "Image", new XAttribute("format", image.Format), ImageSize(img), new XElement(One + "Data", Convert.ToBase64String(image.Data)));
        }

        /// <summary>
        /// Emits the block content of <paramref name="nodes"/> as OEs under an OEChildren.
        /// </summary>
        public void Blocks(IEnumerable<INode> nodes, XElement parent)
        {
            var run = new List<INode>();

            void Flush()
            {
                // A line break between blocks isn't a line of its own
                if (run.All(n => n is IElement { LocalName: "br" } || (n is IText t && string.IsNullOrWhiteSpace(t.Data))))
                {
                    run.Clear();
                    return;
                }

                var text = InlineHtml(run, inTable: _inTable && fromMarkdown, imageLinked: CountLink).Trim();
                if (text.Length > 0)
                {
                    Oe(parent, text, style: _inQuote ? StyleIndex("blockquote") : null);
                }

                run.Clear();
            }

            foreach (var n in nodes)
            {
                // OneNote holds an image as its own OE, so it ends the text before it
                if (EmbeddedImage(n) is { } image)
                {
                    Flush();
                    Oe(parent, "", style: _inQuote ? StyleIndex("blockquote") : null, content: image);
                    continue;
                }

                if (IsInlineNode(n))
                {
                    run.Add(n);
                    continue;
                }

                // Skipped without a flush, so a script mid-paragraph doesn't split it
                if (n is not IElement e || NonContent.Contains(e.LocalName))
                {
                    continue;
                }

                Flush();
                Block(e, parent);
            }

            Flush();
        }

        private void Block(IElement e, XElement parent)
        {
            switch (e.LocalName)
            {
                case "br":
                    break;
                case "h1" or "h2" or "h3" or "h4" or "h5" or "h6":
                    // From Markdown, a blank line before a heading, like hand-made pages. HTML says where its blank lines go.
                    if (fromMarkdown && parent.HasElements && !IsBlank(parent.Elements(One + "OE").LastOrDefault()))
                    {
                        Oe(parent, "");
                    }

                    var headingTags = TakeTags(e);
                    var heading = Oe(parent, InlineHtml(e.ChildNodes, imageLinked: CountLink), style: StyleIndex(e.LocalName));
                    ParagraphFormat(heading, e);
                    AddTags(heading, headingTags);
                    break;
                case "blockquote":
                    Quote(e, parent);
                    break;
                case "ul" or "ol":
                    List(e, parent);
                    break;
                case "table":
                    // Tables can't hold the quote style
                    var tableInQuote = _inQuote;
                    var outerTable = _inTable;
                    (_inQuote, _inTable) = (false, true);
                    Table(e, parent);
                    (_inQuote, _inTable) = (tableInQuote, outerTable);
                    break;
                case "pre":
                    Code(e, parent);
                    break;
                case "hr":
                    Oe(parent, new string('―', 20));
                    break;
                default: // p, div, section, etc.
                    Paragraph(e, parent);
                    break;
            }
        }

        /// <summary>A paragraph or container, with its alignment and OE style.</summary>
        /// <remarks>
        /// - An empty <c>p</c>, or one with only <c>br</c>, is a blank line, as onenote_get_page's HTML writes one.
        /// - A <c>div class="indent"</c> nests its content under the OE before it, as OneNote indents.
        /// </remarks>
        private void Paragraph(IElement e, XElement parent)
        {
            if (e.LocalName == "div" && e.ClassList.Contains("indent") && parent.Elements(One + "OE").LastOrDefault() is { } owner)
            {
                // An empty indent adds nothing: OneNote may reject an empty OEChildren
                var children = owner.Element(One + "OEChildren") ?? new XElement(One + "OEChildren");
                Blocks(e.ChildNodes, children);
                if (children.Parent is null && children.HasElements)
                {
                    owner.Add(children);
                }

                return;
            }

            var tags = TakeTags(e);
            var before = parent.Elements(One + "OE").Count();
            Blocks(e.ChildNodes, parent);
            var added = parent.Elements(One + "OE").Skip(before).ToList();
            if (added.Count == 0 && e.LocalName == "p" && e.Children.All(c => c.LocalName == "br") && string.IsNullOrWhiteSpace(e.TextContent))
            {
                added.Add(Oe(parent, "", style: _inQuote ? StyleIndex("blockquote") : null));
            }

            foreach (var oe in added)
            {
                ParagraphFormat(oe, e);
            }

            AddTags(added.FirstOrDefault(), tags);
        }

        /// <summary>Copies a block element's font, color, and alignment to its OE.</summary>
        private static void ParagraphFormat(XElement oe, IElement e)
        {
            var style = e.GetAttribute("style");
            if (Css.Paragraph(style) is { Length: > 0 } css && oe.Attribute("style") is null)
            {
                oe.SetAttributeValue("style", css);
            }

            // An inner paragraph's alignment wins over its container's
            var alignment = Css.Value(style, "text-align")?.ToLowerInvariant();
            if (alignment is "center" or "right" or "left" && oe.Attribute("alignment") is null)
            {
                oe.SetAttributeValue("alignment", alignment);
            }
        }

        private static bool IsBlank(XElement? oe) =>
            oe is not null && oe.Elements().All(c => c.Name == One + "T") && string.IsNullOrWhiteSpace(string.Concat(oe.Elements(One + "T").Select(t => t.Value)));

        private void Quote(IElement quote, XElement parent)
        {
            if (!_inQuote && HasStyle(parent.Elements(One + "OE").LastOrDefault(), "blockquote"))
            {
                Oe(parent, ""); // keeps back-to-back quotes apart
            }

            var outerQuote = _inQuote;
            _inQuote = true;
            Blocks(quote.ChildNodes, parent);
            _inQuote = outerQuote;
        }

        private void List(IElement list, XElement parent)
        {
            // Keeps back-to-back lists of the same kind apart. Bullets then numbers, or a list then a quoted list, already read as two lists.
            var kind = list.LocalName == "ol" ? "Number" : "Bullet";
            var previous = parent.Elements(One + "OE").LastOrDefault();
            if (previous?.Element(One + "List")?.Element(One + kind) is not null && HasStyle(previous, "blockquote") == _inQuote)
            {
                Oe(parent, "", style: _inQuote ? StyleIndex("blockquote") : null);
            }

            ListItems(list, parent);
        }

        private void Code(IElement pre, XElement parent)
        {
            var code = StyleIndex("code");
            // Keeps back-to-back code blocks apart. An indent under the last code line means one block, split to show the indent.
            if (parent.Elements(One + "OE").LastOrDefault() is { } last && HasStyle(last, "code") && last.Element(One + "OEChildren") is null)
            {
                Oe(parent, "");
            }

            // Only the final newline ends the code. Leading and inner blank lines are code.
            var codeText = pre.TextContent.EndsWith('\n') ? pre.TextContent[..^1] : pre.TextContent;
            foreach (var line in codeText.Split('\n'))
            {
                var escaped = CodeHtml(line.TrimEnd('\r'));
                Oe(parent, escaped.Length > 0 ? escaped : "&nbsp;", style: code);
            }
        }

        private bool HasStyle(XElement? oe, string style) =>
            oe is not null && indexes.TryGetValue(style, out var index) && (string?)oe.Attribute("quickStyleIndex") == index.ToString();

        private void ListItems(IElement list, XElement parent)
        {
            var number = int.TryParse(list.GetAttribute("start"), out var start) ? start : 1;
            var format = "##.";
            if (list.LocalName == "ol")
            {
                var alt = _altNumberFormat.GetValueOrDefault(parent);
                _altNumberFormat[parent] = !alt;
                format = alt ? "##)" : "##.";
            }

            foreach (var li in list.Children.Where(c => c.LocalName == "li"))
            {
                ListItem(li, parent, list.LocalName, (number++, format));
            }
        }

        /// <summary>Emits one list item, with its images and nested blocks as children.</summary>
        /// <remarks>
        /// - An item that's only an image holds the image in place of text. Other images in the item text go below it.
        /// </remarks>
        private void ListItem(IElement li, XElement parent, string listKind, (int Value, string Format) number)
        {
            // A loose item's checkbox sits in its first paragraph, with no item text before it
            var tags = TakeTags(li);
            if (tags.Count == 0 && li.FirstElementChild is { LocalName: "p" } first && string.IsNullOrWhiteSpace(InlineHtml(li.ChildNodes.Where(IsInlineNode))))
            {
                tags = TakeTags(first);
            }

            // A Markdown task item in a bulleted list is a checkbox paragraph, the usual OneNote To Do
            string? kind = fromMarkdown && listKind == "ul" && tags.Any(t => t.IsCheckbox) ? null : listKind;
            var inline = li.ChildNodes.Where(IsInlineNode).ToList();
            var nested = li.ChildNodes.Where(c => !IsInlineNode(c)).ToList();

            // Loose lists wrap the item text in <p>: use the first paragraph as the item text.
            if (string.IsNullOrWhiteSpace(InlineHtml(inline)) && nested.FirstOrDefault() is IElement { LocalName: "p" } p)
            {
                inline = [.. p.ChildNodes];
                nested.RemoveAt(0);
            }

            var images = inline.Select(n => (Node: n, Image: EmbeddedImage(n))).Where(x => x.Image is not null).ToList();
            inline = [.. inline.Except(images.Select(x => x.Node))];
            var imageOnly = images.Count > 0 && string.IsNullOrWhiteSpace(InlineHtml(inline));
            var style = _inQuote ? StyleIndex("blockquote") : (int?)null;
            var oe = imageOnly
                ? Oe(parent, "", style: style, listKind: kind, number: number, content: images[0].Image)
                : Oe(parent, InlineHtml(inline, imageLinked: CountLink), style: style, listKind: kind, number: number);
            ParagraphFormat(oe, li);
            AddTags(oe, tags);

            var children = new XElement(One + "OEChildren");
            foreach (var (_, image) in images.Skip(imageOnly ? 1 : 0))
            {
                Oe(children, "", style: style, content: image);
            }

            Blocks(nested, children);
            if (children.HasElements)
            {
                oe.Add(children);
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

            // From HTML, a first row of only th is OneNote's own header row. Other header cells, and all from Markdown, are bold text.
            var headerRow = !fromMarkdown && rows[0].Count > 0 && rows[0].All(c => c.LocalName == "th");
            if (headerRow)
            {
                tbl.SetAttributeValue("hasHeaderRow", "true");
            }

            foreach (var (row, rowIndex) in rows.Select((row, rowIndex) => (row, rowIndex)))
            {
                var r = new XElement(One + "Row");
                for (var i = 0; i < columns; i++)
                {
                    var cellChildren = new XElement(One + "OEChildren");
                    if (i < row.Count)
                    {
                        if (row[i].LocalName == "th" && !(headerRow && rowIndex == 0))
                        {
                            Oe(cellChildren, $"<b>{InlineHtml(row[i].ChildNodes, inTable: fromMarkdown, imageLinked: CountLink).Trim()}</b>");
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

                    var cell = new XElement(One + "Cell", cellChildren);
                    if (i < row.Count && Shading(row[i]) is { } shading)
                    {
                        cell.SetAttributeValue("shadingColor", shading);
                    }

                    r.Add(cell);
                }

                tbl.Add(r);
            }

            parent.Add(new XElement(One + "OE", tbl));
        }
    }

    /// <summary>A cell's background color, as OneNote's <c>#RRGGBB</c> shading. Null for none or another form.</summary>
    private static string? Shading(IElement cell) =>
        (Css.Value(cell.GetAttribute("style"), "background-color") ?? Css.Value(cell.GetAttribute("style"), "background")) is { } color
        && Css.IsHexColor(color) ? color : null;

    /// <summary>An image's size from its <c>style</c>, in points, as onenote_get_page's HTML writes it.</summary>
    private static XElement? ImageSize(IElement img) =>
        Points(Css.Value(img.GetAttribute("style"), "width")) is { } width && Points(Css.Value(img.GetAttribute("style"), "height")) is { } height
            ? new XElement(One + "Size", new XAttribute("width", width), new XAttribute("height", height), new XAttribute("isSetByUser", "true"))
            : null;

    /// <summary>A size in points, above zero, as OneNote writes it. Null for anything else.</summary>
    private static string? Points(string? value) => ParsePoints(value) is { } points && points > 0 ? Format(points) : null;

    /// <summary>A position in points, which can be zero or negative. Null for anything else.</summary>
    private static string? Coordinate(string? value) => ParsePoints(value) is { } points ? Format(points) : null;

    /// <remarks>
    /// - Finite and within a million points: OneNote rejects a page with <c>NaN</c> or a huge coordinate.
    /// </remarks>
    private static double? ParsePoints(string? value) =>
        value is not null && value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
        && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var points)
        && double.IsFinite(points) && Math.Abs(points) <= 1_000_000
            ? Math.Round(points, 2) + 0.0 // rounded as written, and -0 made 0
            : null;

    private static string Format(double points) => points.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

    // Not \s, which would turn a non-breaking space into a plain one
    [GeneratedRegex(@"[ \t\r\n\f]+")]
    private static partial Regex Whitespace();
}
