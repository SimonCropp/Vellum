namespace Vellum.DocumentBuilder;

/// <summary>
/// Builds DOCX document elements.
/// </summary>
public interface IDocxBuilder : IDisposable
{
    /// <summary>
    /// Adds a heading to the document.
    /// </summary>
    void AddHeading(string text, int level);

    /// <summary>
    /// Adds a paragraph to the document.
    /// </summary>
    void AddParagraph(string text);

    /// <summary>
    /// Starts a new paragraph that can contain inline elements.
    /// </summary>
    void StartParagraph();

    /// <summary>
    /// Ends the current paragraph.
    /// </summary>
    void EndParagraph();

    /// <summary>
    /// Adds inline text to the current paragraph.
    /// </summary>
    void AddText(string text);

    /// <summary>
    /// Adds bold text to the current paragraph.
    /// </summary>
    void AddBoldText(string text);

    /// <summary>
    /// Adds italic text to the current paragraph.
    /// </summary>
    void AddItalicText(string text);

    /// <summary>
    /// Adds inline code to the current paragraph.
    /// </summary>
    void AddInlineCode(string code);

    /// <summary>
    /// Adds a hyperlink to the current paragraph.
    /// </summary>
    void AddHyperlink(string text, string url);

    /// <summary>
    /// Adds a code block to the document.
    /// </summary>
    void AddCodeBlock(string code, string? language = null);

    /// <summary>
    /// Adds a blockquote to the document.
    /// </summary>
    void AddBlockQuote(string text);

    /// <summary>
    /// Starts an unordered list.
    /// </summary>
    void StartUnorderedList();

    /// <summary>
    /// Starts an ordered list.
    /// </summary>
    void StartOrderedList();

    /// <summary>
    /// Ends the current list.
    /// </summary>
    void EndList();

    /// <summary>
    /// Adds a list item.
    /// </summary>
    void AddListItem(string text);

    /// <summary>
    /// Starts a list item that can contain inline elements.
    /// </summary>
    void StartListItem();

    /// <summary>
    /// Ends the current list item.
    /// </summary>
    void EndListItem();

    /// <summary>
    /// Adds a horizontal rule to the document.
    /// </summary>
    void AddHorizontalRule();

    /// <summary>
    /// Adds an HTML chunk using AltChunk.
    /// </summary>
    void AddHtmlChunk(string html);

    /// <summary>
    /// Adds an image to the document.
    /// </summary>
    void AddImage(string altText, string url);

    /// <summary>
    /// Adds a line break within the current paragraph.
    /// </summary>
    void AddLineBreak();

    /// <summary>
    /// Saves the document.
    /// </summary>
    void Save();
}
