using System.Text;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Vellum.DocumentBuilder;

namespace Vellum.Rendering;

/// <summary>
/// Handles HTML content in markdown by converting it to DOCX AltChunks.
/// </summary>
public sealed class HtmlAltChunkHandler
{
    private readonly IDocxBuilder _builder;
    private readonly StringBuilder _htmlBuffer = new();
    private bool _isBuffering;

    public HtmlAltChunkHandler(IDocxBuilder builder)
    {
        _builder = builder;
    }

    /// <summary>
    /// Handles an HTML block from the markdown AST.
    /// </summary>
    public void HandleHtmlBlock(HtmlBlock htmlBlock)
    {
        var html = GetHtmlBlockContent(htmlBlock);
        if (!string.IsNullOrWhiteSpace(html))
        {
            _builder.AddHtmlChunk(html);
        }
    }

    /// <summary>
    /// Handles inline HTML from the markdown AST.
    /// </summary>
    public void HandleHtmlInline(HtmlInline htmlInline)
    {
        var tag = htmlInline.Tag;

        // Detect if this is a self-closing tag or we need to buffer
        if (IsSelfClosingTag(tag))
        {
            _builder.AddHtmlChunk(tag);
        }
        else if (IsOpeningTag(tag))
        {
            StartBuffering();
            _htmlBuffer.Append(tag);
        }
        else if (IsClosingTag(tag))
        {
            _htmlBuffer.Append(tag);
            FlushBuffer();
        }
        else
        {
            // Raw HTML content
            if (_isBuffering)
            {
                _htmlBuffer.Append(tag);
            }
            else
            {
                _builder.AddHtmlChunk(tag);
            }
        }
    }

    /// <summary>
    /// Appends text content to the HTML buffer if buffering is active.
    /// </summary>
    public bool TryBufferText(string text)
    {
        if (_isBuffering)
        {
            _htmlBuffer.Append(System.Net.WebUtility.HtmlEncode(text));
            return true;
        }
        return false;
    }

    private void StartBuffering()
    {
        _isBuffering = true;
        _htmlBuffer.Clear();
    }

    private void FlushBuffer()
    {
        if (_htmlBuffer.Length > 0)
        {
            _builder.AddHtmlChunk(_htmlBuffer.ToString());
        }
        _htmlBuffer.Clear();
        _isBuffering = false;
    }

    private static string GetHtmlBlockContent(HtmlBlock htmlBlock)
    {
        var lines = htmlBlock.Lines;
        var sb = new StringBuilder();
        for (var i = 0; i < lines.Count; i++)
        {
            if (i > 0) sb.AppendLine();
            sb.Append(lines.Lines[i].Slice.ToString());
        }
        return sb.ToString();
    }

    private static bool IsSelfClosingTag(string tag)
    {
        var lowerTag = tag.ToLowerInvariant();
        return lowerTag.Contains("<br") ||
               lowerTag.Contains("<hr") ||
               lowerTag.Contains("<img") ||
               lowerTag.Contains("<input") ||
               lowerTag.Contains("<meta") ||
               lowerTag.Contains("<link") ||
               tag.TrimEnd().EndsWith("/>");
    }

    private static bool IsOpeningTag(string tag)
    {
        return tag.StartsWith('<') &&
               !tag.StartsWith("</") &&
               !tag.TrimEnd().EndsWith("/>");
    }

    private static bool IsClosingTag(string tag)
    {
        return tag.StartsWith("</");
    }
}
