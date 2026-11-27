using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace OneNoteMcp;

/// <summary>
/// Converts Markdown to the HTML that <see cref="HtmlToOneNote"/> takes.
/// </summary>
public static class MarkdownToOneNote
{
    // Strikethrough only: the other extras make ~, ^, ++, and == into formatting.
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .UseEmphasisExtras(Markdig.Extensions.EmphasisExtras.EmphasisExtraOptions.Strikethrough)
        .Build();

    public static string ToHtml(string markdown) => Markdown.ToHtml(markdown, Pipeline);

    /// <summary>
    /// Splits a leading level-1 heading (<c># Title</c> or setext) off the Markdown. Returns a null title when there isn't one.
    /// </summary>
    public static (string? Title, string Body) SplitTitle(string markdown)
    {
        markdown = markdown.TrimStart('﻿');
        if (Markdown.Parse(markdown, Pipeline).FirstOrDefault() is not HeadingBlock { Level: 1, Inline: { } inline } heading)
        {
            return (null, markdown);
        }

        var title = PlainText(inline).Trim();
        // The body starts on the line after the heading. The span can stop before a closing "#".
        var lineEnd = markdown.IndexOf('\n', heading.Span.End);
        return (title, lineEnd < 0 ? "" : markdown[(lineEnd + 1)..]);
    }

    /// <summary>
    /// The heading's text without formatting: literals, code, entities, and a space per line break.
    /// </summary>
    private static string PlainText(ContainerInline container) => string.Concat(container.Select(inline => inline switch
    {
        LiteralInline literal => literal.Content.ToString(),
        CodeInline code => code.Content,
        HtmlEntityInline entity => entity.Transcoded.ToString(),
        LineBreakInline => " ",
        ContainerInline child => PlainText(child),
        _ => "",
    }));
}
