using System.Dynamic;
using System.Text.Json;
using CliFx;
using CliFx.Attributes;
using CliFx.Infrastructure;

namespace Vellum.Cli;

[Command(Description = "Converts a Markdown template to a DOCX document.")]
public class ConvertCommand : ICommand
{
    [CommandParameter(0, Name = "input", Description = "Path to the input Markdown file.")]
    public required string InputPath { get; init; }

    [CommandParameter(1, Name = "output", Description = "Path to the output DOCX file.")]
    public required string OutputPath { get; init; }

    [CommandOption("data", 'd', Description = "Path to a JSON file containing the data model.")]
    public string? DataPath { get; init; }

    public async ValueTask ExecuteAsync(IConsole console)
    {
        // Load the data model
        ExpandoObject model;
        if (!string.IsNullOrEmpty(DataPath))
        {
            var json = await File.ReadAllTextAsync(DataPath);
            model = JsonSerializer.Deserialize<ExpandoObject>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            }) ?? new ExpandoObject();
        }
        else
        {
            model = new ExpandoObject();
        }

        // Convert the markdown to DOCX
        var converter = new MarkdownToDocxConverter();

        await using var inputStream = File.OpenRead(InputPath);
        await using var outputStream = File.Create(OutputPath);

        await converter.ConvertAsync(inputStream, model, outputStream);

        await console.Output.WriteLineAsync($"Successfully converted '{InputPath}' to '{OutputPath}'");
    }
}
