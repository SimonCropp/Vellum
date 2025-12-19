using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Vellum.DocumentBuilder;

namespace Vellum.Rendering;

/// <summary>
/// Renders a Markdig AST to a DOCX document.
/// </summary>
public sealed class MarkdownRenderer
{
    private readonly IDocxBuilder _builder;
    private readonly HtmlAltChunkHandler _htmlHandler;

    public MarkdownRenderer(IDocxBuilder builder)
    {
        _builder = builder;
        _htmlHandler = new HtmlAltChunkHandler(builder);
    }

    /// <summary>
    /// Renders the markdown document to DOCX.
    /// </summary>
    public void Render(MarkdownDocument document)
    {
        foreach (var block in document)
        {
            RenderBlock(block);
        }
    }

    private void RenderBlock(Block block)
    {
        switch (block)
        {
            case HeadingBlock heading:
                RenderHeading(heading);
                break;

            case ParagraphBlock paragraph:
                RenderParagraph(paragraph);
                break;

            case FencedCodeBlock fencedCode:
                RenderFencedCodeBlock(fencedCode);
                break;

            case CodeBlock code:
                RenderCodeBlock(code);
                break;

            case QuoteBlock quote:
                RenderQuoteBlock(quote);
                break;

            case ListBlock list:
                RenderList(list);
                break;

            case Table table:
                RenderTable(table);
                break;

            case ThematicBreakBlock:
                _builder.AddHorizontalRule();
                break;

            case HtmlBlock html:
                _htmlHandler.HandleHtmlBlock(html);
                break;

            case ContainerBlock container:
                foreach (var child in container)
                {
                    RenderBlock(child);
                }
                break;
        }
    }

    private void RenderHeading(HeadingBlock heading)
    {
        var text = GetInlineText(heading.Inline);
        _builder.AddHeading(text, heading.Level);
    }

    private void RenderParagraph(ParagraphBlock paragraph)
    {
        if (paragraph.Inline == null) return;

        _builder.StartParagraph();
        RenderInlines(paragraph.Inline);
        _builder.EndParagraph();
    }

    private void RenderInlines(ContainerInline? container)
    {
        if (container == null) return;

        foreach (var inline in container)
        {
            RenderInline(inline);
        }
    }

    private void RenderInline(Inline inline)
    {
        switch (inline)
        {
            case LiteralInline literal:
                _builder.AddText(literal.Content.ToString());
                break;

            case EmphasisInline emphasis:
                RenderEmphasis(emphasis);
                break;

            case CodeInline code:
                _builder.AddInlineCode(code.Content);
                break;

            case LinkInline link:
                RenderLink(link);
                break;

            case LineBreakInline lineBreak:
                if (lineBreak.IsHard)
                {
                    _builder.AddLineBreak();
                }
                else
                {
                    _builder.AddText(" ");
                }
                break;

            case HtmlInline html:
                _htmlHandler.HandleHtmlInline(html);
                break;

            case AutolinkInline autolink:
                _builder.AddHyperlink(autolink.Url, autolink.Url);
                break;

            case ContainerInline container:
                RenderInlines(container);
                break;
        }
    }

    private void RenderEmphasis(EmphasisInline emphasis)
    {
        // Collect all text from the emphasis
        var text = GetContainerInlineText(emphasis);

        if (emphasis.DelimiterCount == 2)
        {
            _builder.AddBoldText(text);
        }
        else
        {
            _builder.AddItalicText(text);
        }
    }

    private void RenderLink(LinkInline link)
    {
        if (link.IsImage)
        {
            var altText = GetContainerInlineText(link);
            _builder.AddImage(altText, link.Url ?? string.Empty);
        }
        else
        {
            var text = GetContainerInlineText(link);
            _builder.AddHyperlink(text, link.Url ?? string.Empty);
        }
    }

    private void RenderFencedCodeBlock(FencedCodeBlock fencedCode)
    {
        var code = GetCodeBlockText(fencedCode);
        _builder.AddCodeBlock(code, fencedCode.Info);
    }

    private void RenderCodeBlock(CodeBlock code)
    {
        var text = GetCodeBlockText(code);
        _builder.AddCodeBlock(text);
    }

    private void RenderQuoteBlock(QuoteBlock quote)
    {
        // Render each block in the quote
        foreach (var block in quote)
        {
            if (block is ParagraphBlock paragraph)
            {
                var text = GetInlineText(paragraph.Inline);
                _builder.AddBlockQuote(text);
            }
            else
            {
                RenderBlock(block);
            }
        }
    }

    private void RenderList(ListBlock list)
    {
        if (list.IsOrdered)
        {
            _builder.StartOrderedList();
        }
        else
        {
            _builder.StartUnorderedList();
        }

        foreach (var item in list)
        {
            if (item is ListItemBlock listItem)
            {
                RenderListItem(listItem);
            }
        }

        _builder.EndList();
    }

    private void RenderTable(Table table)
    {
        // Count columns from the first row
        var columnCount = 0;
        if (table.Count > 0 && table[0] is TableRow firstRow)
        {
            columnCount = firstRow.Count;
        }

        if (columnCount == 0) return;

        _builder.StartTable(columnCount);

        var isFirstRow = true;
        foreach (var block in table)
        {
            if (block is TableRow row)
            {
                RenderTableRow(row, isFirstRow);
                isFirstRow = false;
            }
        }

        _builder.EndTable();
    }

    private void RenderTableRow(TableRow row, bool isHeader)
    {
        _builder.StartTableRow(isHeader);

        foreach (var block in row)
        {
            if (block is TableCell cell)
            {
                RenderTableCell(cell);
            }
        }

        _builder.EndTableRow();
    }

    private void RenderTableCell(TableCell cell)
    {
        _builder.StartTableCell();

        foreach (var block in cell)
        {
            if (block is ParagraphBlock paragraph)
            {
                _builder.StartParagraph();
                RenderInlines(paragraph.Inline);
                _builder.EndParagraph();
            }
            else
            {
                RenderBlock(block);
            }
        }

        _builder.EndTableCell();
    }

    private void RenderListItem(ListItemBlock listItem)
    {
        _builder.StartListItem();

        foreach (var block in listItem)
        {
            if (block is ParagraphBlock paragraph)
            {
                RenderInlines(paragraph.Inline);
            }
            else if (block is ListBlock nestedList)
            {
                _builder.EndListItem();
                RenderList(nestedList);
                return;
            }
            else
            {
                RenderBlock(block);
            }
        }

        _builder.EndListItem();
    }

    private static string GetInlineText(ContainerInline? container)
    {
        if (container == null) return string.Empty;
        return GetContainerInlineText(container);
    }

    private static string GetContainerInlineText(ContainerInline container)
    {
        var text = new System.Text.StringBuilder();
        foreach (var inline in container)
        {
            text.Append(GetInlineTextRecursive(inline));
        }
        return text.ToString();
    }

    private static string GetInlineTextRecursive(Inline inline)
    {
        return inline switch
        {
            LiteralInline literal => literal.Content.ToString(),
            ContainerInline container => GetContainerInlineText(container),
            CodeInline code => code.Content,
            _ => string.Empty
        };
    }

    private static string GetCodeBlockText(CodeBlock codeBlock)
    {
        var lines = codeBlock.Lines;
        var text = new System.Text.StringBuilder();
        for (var i = 0; i < lines.Count; i++)
        {
            if (i > 0) text.AppendLine();
            text.Append(lines.Lines[i].Slice.ToString());
        }
        return text.ToString();
    }
}
