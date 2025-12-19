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

    [Test]
    public async Task ConvertMarkdownWithTable_CreatesValidDocx()
    {
        // Arrange
        var markdown = """
            # Product Catalog

            | Product | Price | Quantity |
            |---------|-------|----------|
            | Apple   | $1.00 | 100      |
            | Banana  | $0.50 | 150      |
            | Orange  | $0.75 | 200      |
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
    public async Task ConvertMarkdownWithLiquidLoop_CreatesValidDocx()
    {
        // Arrange
        var markdown = """
            # {{ title }}

            ## Team Members

            {% for person in people %}
            - **{{ person.name }}** - {{ person.role }}
            {% endfor %}

            ## Summary

            Total members: {{ people.size }}
            """;
        var model = new
        {
            title = "Project Team",
            people = new[]
            {
                new { name = "Alice", role = "Developer" },
                new { name = "Bob", role = "Designer" },
                new { name = "Charlie", role = "Manager" }
            }
        };
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
    public async Task ConvertMarkdownWithLiquidConditionals_CreatesValidDocx()
    {
        // Arrange
        var markdown = """
            # Order Confirmation

            **Order #{{ order.id }}**

            {% if order.is_priority %}
            > **PRIORITY ORDER** - Expedited shipping enabled
            {% endif %}

            ## Items

            {% for item in order.items %}
            - {{ item.name }} (x{{ item.quantity }}){% if item.on_sale %} *SALE*{% endif %}
            {% endfor %}

            ## Shipping

            {% if order.is_international %}
            International shipping to **{{ order.country }}**

            *Customs documentation will be included.*
            {% else %}
            Domestic shipping

            Estimated delivery: 3-5 business days
            {% endif %}

            {% unless order.items.size == 0 %}
            ---
            Thank you for your order!
            {% endunless %}
            """;
        var model = new
        {
            order = new
            {
                id = 12345,
                is_priority = true,
                is_international = true,
                country = "Canada",
                items = new[]
                {
                    new { name = "Widget", quantity = 2, on_sale = false },
                    new { name = "Gadget", quantity = 1, on_sale = true },
                    new { name = "Gizmo", quantity = 3, on_sale = false }
                }
            }
        };
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
