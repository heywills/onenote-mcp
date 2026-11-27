using System.Net;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static OneNoteMcp.Hierarchy;

namespace OneNoteMcp;

public sealed record ImageRef(string CallbackId, int Index, double? Width, double? Height);

/// <summary>
/// Converts OneNote page XML to Markdown plus image references.
/// </summary>
public static partial class PageMarkdown
{
    public static (string Markdown, IReadOnlyList<ImageRef> Images) Convert(string xml)
    {
        var root = XElement.Parse(xml);
        var title = (string?)root.Attribute("name") ?? (string?)root.Attribute("ID") ?? "Untitled";
        var lines = new List<string> { $"# {title}", "" };
        var images = new List<ImageRef>();

        foreach (var outline in root.Descendants(One + "Outline"))
        {
            Walk(outline, lines, images);
            lines.Add("");
        }

        // Images outside outlines
        foreach (var img in root.Descendants(One + "Image"))
        {
            var id = CallbackId(img);
            if (id is not null && images.All(i => i.CallbackId != id))
            {
                AddImage(img, lines, images);
            }
        }

        return (string.Join("\n", lines).Trim(), images);
    }

    private static void Walk(XElement parent, List<string> lines, List<ImageRef> images)
    {
        foreach (var e in parent.Elements())
        {
            switch (e.Name.LocalName)
            {
                case "T":
                    var text = CleanText(e.Value);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        lines.Add(text);
                    }

                    break;
                case "Image":
                    AddImage(e, lines, images);
                    break;
                case "Table":
                    // Cell text comes from TableRows; list the table's images after it.
                    lines.Add("");
                    lines.AddRange(TableRows(e));
                    lines.Add("");
                    foreach (var img in e.Descendants(One + "Image"))
                    {
                        AddImage(img, lines, images);
                    }

                    continue;
                case "InsertedFile":
                    lines.Add($"[Attached: {(string?)e.Attribute("preferredName") ?? "file"}]");
                    break;
            }

            Walk(e, lines, images);
        }
    }

    private static IEnumerable<string> TableRows(XElement table)
    {
        var rows = table.Elements(One + "Row").Select(row => "| " + string.Join(" | ", row.Elements(One + "Cell").Select(cell =>
            Whitespace().Replace(string.Join(" ", cell.Descendants(One + "T").Select(t => CleanText(t.Value).Trim())), " ").Trim())) + " |").ToList();
        if (rows.Count > 0)
        {
            var columns = rows[0].Count(c => c == '|') - 1;
            rows.Insert(1, "| " + string.Join(" | ", Enumerable.Repeat("---", columns)) + " |");
        }

        return rows;
    }

    private static void AddImage(XElement img, List<string> lines, List<ImageRef> images)
    {
        var id = CallbackId(img);
        if (id is null)
        {
            return;
        }

        var size = img.Element(One + "Size");
        var reference = new ImageRef(id, images.Count + 1, ToDouble(size?.Attribute("width")), ToDouble(size?.Attribute("height")));
        images.Add(reference);
        lines.Add($"[Image {reference.Index}]");
    }

    private static string? CallbackId(XElement img) =>
        (string?)img.Element(One + "CallbackID")?.Attribute("callbackID") ?? (string?)img.Attribute("callbackID");

    private static double? ToDouble(XAttribute? a) =>
        double.TryParse((string?)a, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : null;

    /// <summary>
    /// Strips the inline HTML OneNote keeps inside text elements.
    /// </summary>
    public static string CleanText(string text) =>
        WebUtility.HtmlDecode(Tags().Replace(text, "")).Replace(' ', ' ');

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex Tags();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();
}
