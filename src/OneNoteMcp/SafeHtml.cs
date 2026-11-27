using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace OneNoteMcp;

/// <summary>
/// Keeps the inline HTML that page text can safely carry into a saved HTML file.
/// </summary>
/// <remarks>
/// - OneNote text can hold any HTML, from clipped pages or other tools, and a browser opens the saved file.
/// - Keeps bold, italic, underline, strikethrough, superscript, subscript, line breaks, safe links, and spans with <see cref="Css.Inline"/> styles. Other tags drop and keep their text.
/// </remarks>
internal static class SafeHtml
{
    private static readonly HashSet<string> Plain = ["b", "strong", "i", "em", "u", "s", "strike", "del", "sup", "sub"];

    /// <summary>Elements dropped with their text</summary>
    private static readonly HashSet<string> Dropped = ["script", "style", "template", "noscript", "iframe", "object", "embed", "svg", "math", "title", "textarea"];

    private static readonly HashSet<string> Schemes = new(StringComparer.OrdinalIgnoreCase) { "http", "https", "mailto", "onenote", "file" };

    /// <summary>The safe part of <paramref name="html"/>.</summary>
    public static string Inline(string html)
    {
        var sb = new StringBuilder();
        Write(new HtmlParser().ParseDocument($"<body>{html}</body>").Body!.ChildNodes, sb);
        return sb.ToString();
    }

    /// <summary>Whether a link is relative, a file share, or uses http, https, mailto, onenote, or file.</summary>
    /// <remarks>
    /// - Browsers ignore spaces and control characters in a scheme, so <c>java&#9;script:</c> is still a script.
    /// - A one-letter scheme is a drive letter.
    /// </remarks>
    public static bool IsSafeHref(string href)
    {
        var url = new string([.. href.Where(c => c > ' ')]);
        var colon = url.IndexOf(':');
        var pathStart = url.IndexOfAny(['/', '?', '#']);
        return colon < 0 || (pathStart >= 0 && pathStart < colon) || colon == 1 || Schemes.Contains(url[..colon]);
    }

    private static void Write(INodeList nodes, StringBuilder sb)
    {
        foreach (var n in nodes)
        {
            if (n is IText t)
            {
                sb.Append(Encode(t.Data, attribute: false));
            }
            else if (n is IElement e && !Dropped.Contains(e.LocalName))
            {
                Element(e, sb);
            }
        }
    }

    private static void Element(IElement e, StringBuilder sb)
    {
        var inner = new StringBuilder();
        Write(e.ChildNodes, inner);
        sb.Append(e.LocalName switch
        {
            "br" => "<br>",
            var name when Plain.Contains(name) => $"<{name}>{inner}</{name}>",
            "span" when Css.Inline(e.GetAttribute("style")) is { Length: > 0 } css => $"<span style=\"{Encode(css, attribute: true)}\">{inner}</span>",
            "a" when e.GetAttribute("href") is { } href && IsSafeHref(href) => $"<a href=\"{Encode(href, attribute: true)}\">{inner}</a>",
            _ => inner.ToString(),
        });
    }

    /// <remarks>
    /// - Not <c>WebUtility.HtmlEncode</c>, which turns every accented letter into a number.
    /// - A non-breaking space stays visible as <c>&amp;nbsp;</c>.
    /// </remarks>
    private static string Encode(string text, bool attribute)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            sb.Append(c switch
            {
                '&' => "&amp;",
                '<' => "&lt;",
                '>' => "&gt;",
                '"' when attribute => "&quot;",
                ' ' => "&nbsp;",
                _ => c.ToString(),
            });
        }

        return sb.ToString();
    }
}
