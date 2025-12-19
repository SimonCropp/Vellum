namespace Vellum;

/// <summary>
/// Converts Markdown templates to DOCX documents.
/// </summary>
public interface IMarkdownToDocxConverter
{
    /// <summary>
    /// Converts a Markdown template to a DOCX document.
    /// </summary>
    /// <typeparam name="TModel">The type of the data model.</typeparam>
    /// <param name="markdownStream">The input Markdown template stream.</param>
    /// <param name="model">The data model to merge with the template.</param>
    /// <param name="outputStream">The output stream for the DOCX document.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    Task ConvertAsync<TModel>(
        Stream markdownStream,
        TModel model,
        Stream outputStream,
        CancellationToken cancellationToken = default) where TModel : class;
}
