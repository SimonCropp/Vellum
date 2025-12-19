using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Vellum.DocumentBuilder;

/// <summary>
/// Builds DOCX documents using OpenXML SDK.
/// </summary>
public sealed class DocxBuilder : IDocxBuilder
{
    private readonly WordprocessingDocument _document;
    private readonly Body _body;
    private Paragraph? _currentParagraph;
    private int _altChunkId;
    private readonly Stack<ListContext> _listStack = new();
    private int _abstractNumId;

    // Table state
    private Table? _currentTable;
    private TableRow? _currentTableRow;
    private TableCell? _currentTableCell;
    private int _currentTableColumnCount;
    private bool _isHeaderRow;

    private record ListContext(int AbstractNumId, bool IsOrdered, int Level);

    public DocxBuilder(Stream outputStream)
    {
        _document = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document);
        var mainPart = _document.AddMainDocumentPart();
        mainPart.Document = new Document();
        _body = mainPart.Document.AppendChild(new Body());

        // Add numbering definitions part for lists
        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering();
    }

    public void AddHeading(string text, int level)
    {
        var paragraph = new Paragraph();
        var pPr = new ParagraphProperties();
        pPr.AppendChild(new ParagraphStyleId { Val = $"Heading{Math.Min(level, 9)}" });
        paragraph.AppendChild(pPr);

        var run = new Run();
        run.AppendChild(new Text(text));
        paragraph.AppendChild(run);

        _body.AppendChild(paragraph);
    }

    public void AddParagraph(string text)
    {
        var paragraph = new Paragraph();
        var run = new Run();
        run.AppendChild(new Text(text));
        paragraph.AppendChild(run);
        _body.AppendChild(paragraph);
    }

    public void StartParagraph()
    {
        _currentParagraph = new Paragraph();
        if (_listStack.Count > 0)
        {
            var listContext = _listStack.Peek();
            var pPr = new ParagraphProperties();
            var numPr = new NumberingProperties();
            numPr.AppendChild(new NumberingLevelReference { Val = listContext.Level });
            numPr.AppendChild(new NumberingId { Val = listContext.AbstractNumId + 1 });
            pPr.AppendChild(numPr);
            _currentParagraph.AppendChild(pPr);
        }
    }

    public void EndParagraph()
    {
        if (_currentParagraph != null)
        {
            if (_currentTableCell != null)
            {
                _currentTableCell.AppendChild(_currentParagraph);
            }
            else
            {
                _body.AppendChild(_currentParagraph);
            }
            _currentParagraph = null;
        }
    }

    public void AddText(string text)
    {
        if (_currentParagraph == null) StartParagraph();
        var run = new Run();
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        _currentParagraph!.AppendChild(run);
    }

    public void AddBoldText(string text)
    {
        if (_currentParagraph == null) StartParagraph();
        var run = new Run();
        run.AppendChild(new RunProperties(new Bold()));
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        _currentParagraph!.AppendChild(run);
    }

    public void AddItalicText(string text)
    {
        if (_currentParagraph == null) StartParagraph();
        var run = new Run();
        run.AppendChild(new RunProperties(new Italic()));
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        _currentParagraph!.AppendChild(run);
    }

    public void AddInlineCode(string code)
    {
        if (_currentParagraph == null) StartParagraph();
        var run = new Run();
        var rPr = new RunProperties();
        rPr.AppendChild(new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" });
        rPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "E0E0E0" });
        run.AppendChild(rPr);
        run.AppendChild(new Text(code) { Space = SpaceProcessingModeValues.Preserve });
        _currentParagraph!.AppendChild(run);
    }

    public void AddHyperlink(string text, string url)
    {
        if (_currentParagraph == null) StartParagraph();

        var mainPart = _document.MainDocumentPart!;
        var hyperlinkRelationship = mainPart.AddHyperlinkRelationship(new Uri(url, UriKind.RelativeOrAbsolute), true);

        var hyperlink = new Hyperlink { Id = hyperlinkRelationship.Id };
        var run = new Run();
        var rPr = new RunProperties();
        rPr.AppendChild(new Color { Val = "0000FF" });
        rPr.AppendChild(new Underline { Val = UnderlineValues.Single });
        run.AppendChild(rPr);
        run.AppendChild(new Text(text));
        hyperlink.AppendChild(run);
        _currentParagraph!.AppendChild(hyperlink);
    }

    public void AddCodeBlock(string code, string? language = null)
    {
        var lines = code.Split('\n');
        foreach (var line in lines)
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties();
            pPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "F5F5F5" });
            paragraph.AppendChild(pPr);

            var run = new Run();
            var rPr = new RunProperties();
            rPr.AppendChild(new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" });
            run.AppendChild(rPr);
            run.AppendChild(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
            _body.AppendChild(paragraph);
        }
    }

    public void AddBlockQuote(string text)
    {
        var paragraph = new Paragraph();
        var pPr = new ParagraphProperties();
        pPr.AppendChild(new Indentation { Left = "720" });
        pPr.AppendChild(new ParagraphBorders(
            new LeftBorder { Val = BorderValues.Single, Size = 24, Color = "CCCCCC" }));
        paragraph.AppendChild(pPr);

        var run = new Run();
        var rPr = new RunProperties();
        rPr.AppendChild(new Italic());
        rPr.AppendChild(new Color { Val = "666666" });
        run.AppendChild(rPr);
        run.AppendChild(new Text(text));
        paragraph.AppendChild(run);
        _body.AppendChild(paragraph);
    }

    public void StartUnorderedList()
    {
        var level = _listStack.Count;
        var abstractNumId = CreateNumberingDefinition(isOrdered: false);
        _listStack.Push(new ListContext(abstractNumId, false, level));
    }

    public void StartOrderedList()
    {
        var level = _listStack.Count;
        var abstractNumId = CreateNumberingDefinition(isOrdered: true);
        _listStack.Push(new ListContext(abstractNumId, true, level));
    }

    private int CreateNumberingDefinition(bool isOrdered)
    {
        var numberingPart = _document.MainDocumentPart!.NumberingDefinitionsPart!;
        var numbering = numberingPart.Numbering;

        var abstractNumId = _abstractNumId++;
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };

        for (var i = 0; i < 9; i++)
        {
            var level = new Level { LevelIndex = i };
            level.AppendChild(new StartNumberingValue { Val = 1 });
            level.AppendChild(new NumberingFormat
            {
                Val = isOrdered ? NumberFormatValues.Decimal : NumberFormatValues.Bullet
            });
            level.AppendChild(new LevelText
            {
                Val = isOrdered ? $"%{i + 1}." : "\u2022"
            });
            level.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
            level.AppendChild(new PreviousParagraphProperties(
                new Indentation { Left = ((i + 1) * 720).ToString(), Hanging = "360" }));
            abstractNum.AppendChild(level);
        }

        numbering.InsertAt(abstractNum, 0);

        var numInstance = new NumberingInstance { NumberID = abstractNumId + 1 };
        numInstance.AppendChild(new AbstractNumId { Val = abstractNumId });
        numbering.AppendChild(numInstance);

        return abstractNumId;
    }

    public void EndList()
    {
        if (_listStack.Count > 0)
        {
            _listStack.Pop();
        }
    }

    public void AddListItem(string text)
    {
        StartListItem();
        AddText(text);
        EndListItem();
    }

    public void StartListItem()
    {
        StartParagraph();
    }

    public void EndListItem()
    {
        EndParagraph();
    }

    public void AddHorizontalRule()
    {
        var paragraph = new Paragraph();
        var pPr = new ParagraphProperties();
        pPr.AppendChild(new ParagraphBorders(
            new BottomBorder { Val = BorderValues.Single, Size = 12, Color = "auto" }));
        paragraph.AppendChild(pPr);
        _body.AppendChild(paragraph);
    }

    public void AddHtmlChunk(string html)
    {
        var mainPart = _document.MainDocumentPart!;
        var altChunkId = $"altChunk{++_altChunkId}";

        var chunk = mainPart.AddAlternativeFormatImportPart(
            AlternativeFormatImportPartType.Html, altChunkId);

        using (var stream = chunk.GetStream())
        using (var writer = new StreamWriter(stream))
        {
            // Wrap in basic HTML structure if needed
            var fullHtml = html.Contains("<html", StringComparison.OrdinalIgnoreCase)
                ? html
                : $"<!DOCTYPE html><html><body>{html}</body></html>";
            writer.Write(fullHtml);
        }

        var altChunk = new AltChunk { Id = altChunkId };
        _body.AppendChild(altChunk);
    }

    public void AddImage(string altText, string url)
    {
        // For now, add a placeholder paragraph with the alt text
        // Full image support would require downloading and embedding the image
        var paragraph = new Paragraph();
        var run = new Run();
        run.AppendChild(new Text($"[Image: {altText}]"));
        paragraph.AppendChild(run);
        _body.AppendChild(paragraph);
    }

    public void AddLineBreak()
    {
        if (_currentParagraph == null) StartParagraph();
        var run = new Run();
        run.AppendChild(new Break());
        _currentParagraph!.AppendChild(run);
    }

    public void StartTable(int columnCount)
    {
        _currentTableColumnCount = columnCount;
        _currentTable = new Table();

        // Add table properties with borders
        var tblPr = new TableProperties();

        var tblBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new RightBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" }
        );
        tblPr.AppendChild(tblBorders);

        // Set table width to 100%
        tblPr.AppendChild(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });

        _currentTable.AppendChild(tblPr);

        // Define column widths (equal distribution)
        var tblGrid = new TableGrid();
        var colWidth = 5000 / columnCount;
        for (var i = 0; i < columnCount; i++)
        {
            tblGrid.AppendChild(new GridColumn { Width = colWidth.ToString() });
        }
        _currentTable.AppendChild(tblGrid);
    }

    public void EndTable()
    {
        if (_currentTable != null)
        {
            _body.AppendChild(_currentTable);
            _currentTable = null;
            _currentTableColumnCount = 0;
        }
    }

    public void StartTableRow(bool isHeader = false)
    {
        _isHeaderRow = isHeader;
        _currentTableRow = new TableRow();

        if (isHeader)
        {
            var trPr = new TableRowProperties();
            trPr.AppendChild(new TableHeader());
            _currentTableRow.AppendChild(trPr);
        }
    }

    public void EndTableRow()
    {
        if (_currentTableRow != null && _currentTable != null)
        {
            _currentTable.AppendChild(_currentTableRow);
            _currentTableRow = null;
            _isHeaderRow = false;
        }
    }

    public void StartTableCell()
    {
        _currentTableCell = new TableCell();

        var tcPr = new TableCellProperties();

        // Add shading for header cells
        if (_isHeaderRow)
        {
            tcPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "E0E0E0" });
        }

        // Set vertical alignment to center
        tcPr.AppendChild(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        _currentTableCell.AppendChild(tcPr);
    }

    public void EndTableCell()
    {
        if (_currentTableCell != null && _currentTableRow != null)
        {
            // Ensure cell has at least one paragraph (required by OpenXML)
            if (!_currentTableCell.Elements<Paragraph>().Any())
            {
                _currentTableCell.AppendChild(new Paragraph());
            }

            _currentTableRow.AppendChild(_currentTableCell);
            _currentTableCell = null;
        }
    }

    public void Save()
    {
        _document.Save();
    }

    public void Dispose()
    {
        _document.Dispose();
    }
}
