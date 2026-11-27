namespace OneNoteMcp;

/// <summary>
/// Keeps the CSS OneNote can show from a <c>style</c> attribute, and drops the rest.
/// </summary>
/// <remarks>
/// - Values that could load something or break the markup (<c>url(</c>, <c>expression(</c>, angle brackets, backslashes) are dropped. Callers encode the attribute, so quotes, as in <c>"Lucida Console"</c>, are safe.
/// </remarks>
internal static class Css
{
    private static readonly HashSet<string> InlineProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "color", "background", "background-color", "font-family", "font-size", "font-weight", "font-style", "text-decoration", "vertical-align",
    };

    private static readonly HashSet<string> ParagraphProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "color", "font-family", "font-size", "line-height",
    };

    /// <summary>The text formatting a span may carry.</summary>
    public static string Inline(string? style) => Filter(style, InlineProperties);

    /// <summary>The formatting an OE's own style may carry. Alignment goes in its own attribute.</summary>
    public static string Paragraph(string? style) => Filter(style, ParagraphProperties);

    /// <summary>One property's value, or null.</summary>
    public static string? Value(string? style, string property) =>
        Declarations(style).Where(d => d.Name.Equals(property, StringComparison.OrdinalIgnoreCase)).Select(d => d.Value).LastOrDefault();

    private static string Filter(string? style, HashSet<string> keep) =>
        string.Join(";", Declarations(style).Where(d => keep.Contains(d.Name)).Select(d => $"{d.Name.ToLowerInvariant()}:{d.Value}"));

    private static IEnumerable<(string Name, string Value)> Declarations(string? style) =>
        (style ?? "").Split(';')
            .Select(d => d.Split(':', 2))
            .Where(p => p.Length == 2)
            .Select(p => (Name: p[0].Trim(), Value: string.Join(' ', p[1].Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))))
            .Where(d => d.Name.Length > 0 && d.Value.Length > 0 && IsSafe(d.Value));

    private static bool IsSafe(string value) =>
        value.IndexOfAny(['<', '>', '\\']) < 0
        && !value.Contains("url(", StringComparison.OrdinalIgnoreCase)
        && !value.Contains("expression(", StringComparison.OrdinalIgnoreCase);
}
