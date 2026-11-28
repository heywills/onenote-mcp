using System.Globalization;
using System.Net;
using System.Text;
using System.Xml.Linq;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

/// <summary>
/// Converts OneNote page XML to an HTML document that keeps what Markdown can't: colors, highlights, fonts, alignment, indents, cell shading, and image sizes.
/// </summary>
/// <remarks>
/// - OneNote stores text as HTML, so inline formatting passes through as OneNote wrote it.
/// - Paragraph formatting comes from the OE: its style, alignment, and QuickStyleDef (headings, quote, code).
/// - onenote_create_page reads this HTML back, taking the title from <c>&lt;title&gt;</c>.
/// - Tags, ink, and paragraph spacing aren't kept.
/// </remarks>
public static class PageHtml
{
    /// <param name="imageHtml">The HTML for each image. Default: an <c>img</c> with its callback ID and no source.</param>
    public static (string Html, IReadOnlyList<ImageRef> Images) Convert(string xml, Func<ImageRef, string>? imageHtml = null)
    {
        var root = XElement.Parse(xml);
        var (title, styles, repeatedTitle) = PageMarkdown.Parts(root);
        var writer = new Writer(styles, repeatedTitle, imageHtml ?? (img => ImageTag(img, null)), PageTags.Definitions(root));
        foreach (var outline in root.Descendants(One + "Outline"))
        {
            foreach (var children in outline.Elements(One + "OEChildren"))
            {
                writer.Blocks(children);
            }
        }

        // Images outside outlines
        foreach (var img in root.Descendants(One + "Image").Where(i => i.Ancestors(One + "Outline").FirstOrDefault() is null))
        {
            writer.Image(img);
        }

        var html = $"""
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <title>{Encode(title)}</title>
            </head>
            <body>
            {writer}</body>
            </html>

            """;
        return (html.Replace("\r\n", "\n"), writer.Images);
    }

    /// <summary>An <c>img</c> for a saved image, or with its callback ID when <paramref name="src"/> is null.</summary>
    /// <remarks>
    /// - The size is in points, as OneNote stores it, so the page reads back at the same size.
    /// </remarks>
    public static string ImageTag(ImageRef img, string? src)
    {
        var source = src is null ? $" data-callback-id=\"{Encode(img.CallbackId)}\"" : $" src=\"{Encode(src)}\"";
        var size = img is { Width: { } w, Height: { } h }
            ? $" style=\"width:{w.ToString("0.##", CultureInfo.InvariantCulture)}pt;height:{h.ToString("0.##", CultureInfo.InvariantCulture)}pt\""
            : "";
        return $"<img{source} alt=\"Image {img.Index}\"{size}>";
    }

    /// <summary>An <c>img</c> for an image saved as <c>image-N.extension</c> in <paramref name="folder"/>, or a link for a type browsers can't show.</summary>
    public static string ImageLink(ImageRef img, string folder, string extension)
    {
        var src = $"{Uri.EscapeDataString(folder)}/image-{img.Index}.{extension}";
        return extension == "bin" ? $"<a href=\"{Encode(src)}\">Image {img.Index}: unsupported type</a>" : ImageTag(img, src);
    }

    /// <summary>Text for an image that couldn't be saved, with the reason and the image's text.</summary>
    public static string ImageProblem(int index, string reason, string? text = null) =>
        Encode($"[Image {index}: {PageMarkdown.ProblemText(reason, text)}]");

    private static string Encode(string text) => WebUtility.HtmlEncode(text);

    /// <param name="tagDefinitions">The page's tags by index</param>
    private sealed class Writer(
        Dictionary<string, (string Name, bool Italic)> styles, XElement? repeatedTitle, Func<ImageRef, string> imageHtml, Dictionary<string, Tag> tagDefinitions)
    {
        /// <summary>The OE's tags as a checkbox or label, before its text.</summary>
        private string TagHtml(XElement oe) => PageTags.Html(PageTags.Of(oe, tagDefinitions));

        private readonly StringBuilder _html = new();

        public List<ImageRef> Images { get; } = [];

        /// <summary>Writes an OEChildren's OEs, grouping runs of list items, quotes, and code lines.</summary>
        public void Blocks(XElement oeChildren)
        {
            var oes = oeChildren.Elements(One + "OE").ToList();
            for (var i = 0; i < oes.Count;)
            {
                var kind = Kind(oes[i]);
                var end = i + 1;
                // A code line with content under it ends the code block, which then shows that content
                while (kind is not null && end < oes.Count && Kind(oes[end]) == kind && !(kind == "pre" && HasContent(oes[end - 1])))
                {
                    end++;
                }

                Group(kind, oes.GetRange(i, end - i));
                i = end;
            }
        }

        /// <summary>The element a run of OEs shares: <c>ul</c>, <c>ol</c>, <c>blockquote</c>, or <c>pre</c>. Null for a plain paragraph or heading.</summary>
        private string? Kind(XElement oe) =>
            oe == repeatedTitle ? null
            : oe.Element(One + "List") is { } list ? (list.Element(One + "Number") is null ? "ul" : "ol")
            : StyleName(oe) switch
            {
                "blockquote" => "blockquote",
                "code" => "pre",
                _ => null,
            };

        private static bool HasContent(XElement oe) =>
            oe.Elements().Any(c => c.Name.LocalName is "Image" or "Table" or "InsertedFile" or "OEChildren");

        private string StyleName(XElement oe) =>
            styles.TryGetValue((string?)oe.Attribute("quickStyleIndex") ?? "", out var s) ? s.Name : "";

        private void Group(string? kind, List<XElement> oes)
        {
            switch (kind)
            {
                case "ul" or "ol":
                    _html.Append(kind == "ol" && Start(oes[0]) is { } start and not 1 ? $"<ol start=\"{start}\">\n" : $"<{kind}>\n");
                    foreach (var oe in oes)
                    {
                        _html.Append($"<li{Style(oe)}>{TagHtml(oe)}{Text(oe)}");
                        Content(oe, indent: false);
                        _html.Append("</li>\n");
                    }

                    _html.Append($"</{kind}>\n");
                    break;
                case "blockquote":
                    _html.Append("<blockquote>\n");
                    oes.ForEach(Paragraph);
                    _html.Append("</blockquote>\n");
                    break;
                case "pre":
                    // A line that holds only an image, table, or file isn't a code line
                    var lines = oes.Where(oe => oe.Element(One + "T") is not null || !oe.Elements().Any(c => c.Name.LocalName is "Image" or "Table" or "InsertedFile"))
                        .Select(oe => Encode(PageMarkdown.CleanText(string.Concat(oe.Elements(One + "T").Select(t => t.Value)))))
                        .ToList();
                    if (lines.Count > 0)
                    {
                        // The parser drops the first newline after <pre>, so this one keeps a leading blank line
                        _html.Append($"<pre>\n{string.Join("\n", lines)}</pre>\n");
                    }

                    Content(oes[^1], indent: true);
                    break;
                default:
                    oes.ForEach(Paragraph);
                    break;
            }
        }

        private static int? Start(XElement oe) =>
            (string?)oe.Element(One + "List")?.Element(One + "Number")?.Attribute("text") is { } text
            && int.TryParse(new string([.. text.TakeWhile(char.IsDigit)]), out var n) ? n : null;

        /// <summary>A paragraph or heading, then its image, table, or file, then its indented children.</summary>
        private void Paragraph(XElement oe)
        {
            if (oe != repeatedTitle && oe.Element(One + "T") is not null)
            {
                var tag = StyleName(oe) is { Length: 2 } name && name[0] == 'h' && char.IsDigit(name[1]) ? name : "p";
                // A browser shows an empty <p> with no height
                var text = Text(oe) is var t && tag == "p" && string.IsNullOrWhiteSpace(t) ? "<br>" : t;
                _html.Append($"<{tag}{Style(oe)}>{TagHtml(oe)}{text}</{tag}>\n");
            }

            // The repeated title's children aren't indented under anything shown
            Content(oe, indent: oe != repeatedTitle);
        }

        private static string Text(XElement oe) => SafeHtml.Inline(string.Concat(oe.Elements(One + "T").Select(t => t.Value)));

        /// <summary>The OE's font, color, and alignment, as a <c>style</c> attribute.</summary>
        private static string Style(XElement oe)
        {
            var style = (string?)oe.Attribute("style");
            var css = Css.Paragraph(style);
            // OneNote often copies the alignment into the style
            var alignment = (Css.Value(style, "text-align") ?? (string?)oe.Attribute("alignment"))?.ToLowerInvariant();
            if (alignment is "center" or "right")
            {
                css = css.Length > 0 ? $"{css};text-align:{alignment}" : $"text-align:{alignment}";
            }

            return css.Length > 0 ? $" style=\"{Encode(css)}\"" : "";
        }

        /// <param name="indent">Wrap child OEs in an indent, as OneNote shows them. List items nest them instead.</param>
        private void Content(XElement oe, bool indent)
        {
            foreach (var child in oe.Elements())
            {
                switch (child.Name.LocalName)
                {
                    case "Image":
                        Image(child);
                        break;
                    case "Table":
                        Table(child);
                        break;
                    case "InsertedFile":
                        _html.Append($"<p>[Attached: {Encode((string?)child.Attribute("preferredName") ?? "file")}]</p>\n");
                        break;
                    case "OEChildren" when indent:
                        _html.Append("<div class=\"indent\" style=\"margin-left:2em\">\n");
                        Blocks(child);
                        _html.Append("</div>\n");
                        break;
                    case "OEChildren":
                        Blocks(child);
                        break;
                }
            }
        }

        public void Image(XElement img)
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
                reference = new ImageRef(id, Images.Count + 1, ToDouble(size?.Attribute("width")), ToDouble(size?.Attribute("height")), PageMarkdown.ImageText(img));
                Images.Add(reference);
            }

            _html.Append($"<p>{imageHtml(reference)}</p>\n");
        }

        private static double? ToDouble(XAttribute? a) =>
            double.TryParse((string?)a, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : null;

        /// <remarks>
        /// - The header row is <c>th</c> only when OneNote marks it as one.
        /// - Cell shading becomes the cell's background color.
        /// </remarks>
        private void Table(XElement table)
        {
            var header = (string?)table.Attribute("hasHeaderRow") == "true";
            _html.Append("<table>\n");
            foreach (var (row, r) in table.Elements(One + "Row").Select((row, r) => (row, r)))
            {
                _html.Append("<tr>");
                foreach (var cell in row.Elements(One + "Cell"))
                {
                    var tag = header && r == 0 ? "th" : "td";
                    var shading = (string?)cell.Attribute("shadingColor") is { } color && Css.IsHexColor(color) ?$" style=\"background-color:{Encode(color)}\"" : "";
                    _html.Append($"<{tag}{shading}>\n");
                    foreach (var children in cell.Elements(One + "OEChildren"))
                    {
                        Blocks(children);
                    }

                    _html.Append($"</{tag}>");
                }

                _html.Append("</tr>\n");
            }

            _html.Append("</table>\n");
        }

        public override string ToString() => _html.ToString();
    }
}
