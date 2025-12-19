using System.Text;
using TUnit.Core;
using VerifyTests;

namespace Vellum.Tests;

public class ConverterTests
{
    [Test]
    public async Task ConvertSimpleMarkdown_CreatesValidDocx()
    {
        // Arrange
        var markdown = "# Hello World\n\nThis is a test paragraph.";
        var model = new { };
        var converter = new MarkdownToDocxConverter();

        using var inputStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var outputStream = new MemoryStream();

        // Act
        await converter.ConvertAsync(inputStream, model, outputStream);

        // Assert
        outputStream.Position = 0;
        await Verifier.Verify(outputStream, extension: "docx");
    }

    [Test]
    public async Task ConvertMarkdownWithFormatting_CreatesValidDocx()
    {
        // Arrange
        var markdown = """
            # Formatted Document

            This has **bold** and *italic* text.

            And some `inline code` too.
            """;
        var model = new { };
        var converter = new MarkdownToDocxConverter();

        using var inputStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var outputStream = new MemoryStream();

        // Act
        await converter.ConvertAsync(inputStream, model, outputStream);

        // Assert
        outputStream.Position = 0;
        await Verifier.Verify(outputStream, extension: "docx");
    }

    [Test]
    public async Task ConvertMarkdownWithList_CreatesValidDocx()
    {
        // Arrange
        var markdown = """
            # Shopping List

            - Apples
            - Bananas
            - Oranges

            ## Numbered Steps

            1. First step
            2. Second step
            3. Third step
            """;
        var model = new { };
        var converter = new MarkdownToDocxConverter();

        using var inputStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var outputStream = new MemoryStream();

        // Act
        await converter.ConvertAsync(inputStream, model, outputStream);

        // Assert
        outputStream.Position = 0;
        await Verifier.Verify(outputStream, extension: "docx");
    }

    [Test]
    public async Task ConvertMarkdownWithCodeBlock_CreatesValidDocx()
    {
        // Arrange
        var markdown = """
            # Code Example

            Here is some code:

            ```csharp
            public class Hello
            {
                public void World() => Console.WriteLine("Hello!");
            }
            ```
            """;
        var model = new { };
        var converter = new MarkdownToDocxConverter();

        using var inputStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var outputStream = new MemoryStream();

        // Act
        await converter.ConvertAsync(inputStream, model, outputStream);

        // Assert
        outputStream.Position = 0;
        await Verifier.Verify(outputStream, extension: "docx");
    }

    [Test]
    public async Task ConvertMarkdownWithHtml_UsesAltChunk()
    {
        // Arrange
        var markdown = """
            # Document with HTML

            <div style="color: red;">
                <p>This is HTML content</p>
            </div>
            """;
        var model = new { };
        var converter = new MarkdownToDocxConverter();

        using var inputStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using var outputStream = new MemoryStream();

        // Act
        await converter.ConvertAsync(inputStream, model, outputStream);

        // Assert
        outputStream.Position = 0;
        await Verifier.Verify(outputStream, extension: "docx");
    }

}
