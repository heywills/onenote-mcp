using Markdig;

namespace OneNoteMcp;

/// <summary>
/// Converts Markdown to the HTML that <see cref="HtmlToOneNote"/> takes.
/// </summary>
public static class MarkdownToOneNote
{
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .UseEmphasisExtras()
        .Build();

    public static string ToHtml(string markdown) => Markdown.ToHtml(markdown, Pipeline);

    /// <summary>
    /// Splits a leading <c># Title</c> line off the Markdown. Returns a null title when there isn't one.
    /// </summary>
    public static (string? Title, string Body) SplitTitle(string markdown)
    {
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        var first = Array.FindIndex(lines, l => l.Trim().Length > 0);
        if (first < 0 || !lines[first].StartsWith("# "))
        {
            return (null, markdown);
        }

        return (lines[first][2..].Trim(), string.Join("\n", lines.Skip(first + 1)));
    }
}
