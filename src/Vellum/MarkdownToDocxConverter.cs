using Fluid;
using Markdig;
using Vellum.DocumentBuilder;
using Vellum.Rendering;

namespace Vellum;

/// <summary>
/// Converts Markdown templates to DOCX documents using Liquid templating.
/// </summary>
public sealed class MarkdownToDocxConverter : IMarkdownToDocxConverter
{
    private readonly MarkdownPipeline _pipeline;
    private readonly FluidParser _parser;

    public MarkdownToDocxConverter()
    {
        _pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();

        _parser = new FluidParser();
    }

    /// <inheritdoc />
    public async Task ConvertAsync<TModel>(
        Stream markdownStream,
        TModel model,
        Stream outputStream,
        CancellationToken cancellationToken = default) where TModel : class
    {
        // Step 1: Read the markdown template
        using var reader = new StreamReader(markdownStream);
        var markdownTemplate = await reader.ReadToEndAsync(cancellationToken);

        // Step 2: Process the template with Fluid (Liquid)
        string processedMarkdown;
        if (_parser.TryParse(markdownTemplate, out var template, out var error))
        {
            var context = new TemplateContext(model);
            processedMarkdown = await template.RenderAsync(context);
        }
        else
        {
            throw new InvalidOperationException($"Failed to parse Liquid template: {error}");
        }

        // Step 3: Parse the markdown into an AST
        var document = Markdown.Parse(processedMarkdown, _pipeline);

        // Step 4: Build the DOCX document
        using var docxBuilder = new DocxBuilder(outputStream);
        var renderer = new MarkdownRenderer(docxBuilder);
        renderer.Render(document);
        docxBuilder.Save();
    }
}
