using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace OneNoteMcp;

/// <summary>
/// Converts Markdown to the HTML that <see cref="HtmlToOneNote"/> takes.
/// </summary>
public static class MarkdownToOneNote
{
    private const char ByteOrderMark = '\uFEFF';

    /// <summary>Pipe tables and strikethrough.</summary>
    /// <remarks>The other emphasis extras would make ~, ^, ++, and == into formatting.</remarks>
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .UseEmphasisExtras(Markdig.Extensions.EmphasisExtras.EmphasisExtraOptions.Strikethrough)
        .Build();

    public static string ToHtml(string markdown) => Markdown.ToHtml(markdown.TrimStart(ByteOrderMark), Pipeline);

    /// <summary>
    /// Splits a leading level-1 heading (<c># Title</c> or setext) off the Markdown. Returns a null title when there isn't one.
    /// </summary>
    /// <remarks>HTML comments and link reference definitions before the heading don't count, and stay in the body.</remarks>
    public static (string? Title, string Body) SplitTitle(string markdown)
    {
        markdown = markdown.TrimStart(ByteOrderMark);
        var first = Markdown.Parse(markdown, Pipeline).FirstOrDefault(b => b is not (HtmlBlock or LinkReferenceDefinitionGroup));
        if (first is not HeadingBlock { Level: 1, Inline: { } inline } heading)
        {
            return (null, markdown);
        }

        var title = PlainText(inline).Trim();
        if (title.Length == 0)
        {
            return (null, markdown);
        }

        // The body is what's before the heading plus what starts on the line after it. The span can stop before a closing "#".
        var lineEnd = markdown.IndexOf('\n', heading.Span.End);
        return (title, markdown[..heading.Span.Start] + (lineEnd < 0 ? "" : markdown[(lineEnd + 1)..]));
    }

    /// <summary>
    /// The heading's text without formatting: literals, code, entities, autolinks, and a space per line break.
    /// </summary>
    private static string PlainText(ContainerInline container) => string.Concat(container.Select(inline => inline switch
    {
        LiteralInline literal => literal.Content.ToString(),
        CodeInline code => code.Content,
        HtmlEntityInline entity => entity.Transcoded.ToString(),
        AutolinkInline autolink => autolink.Url,
        LineBreakInline => " ",
        ContainerInline child => PlainText(child),
        _ => "",
    }));
}
