using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Text.Json;

namespace ExcelForge.Converters;

/// <summary>
/// Exports an Excel worksheet to JSON.
/// The first row is treated as the header / field names.
/// Output is an array of objects: [{ "Column A": "val", "Column B": "val" }, ...]
/// </summary>
internal sealed class JsonExporter : IConverter
{
    private readonly ILogger<JsonExporter> _logger;

    public JsonExporter(ILogger<JsonExporter> logger)
    {
        _logger = logger;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task ConvertAsync(string sourcePath, string destinationPath,
                                   CancellationToken ct = default)
    {
        _logger.LogInformation("JsonExporter: reading {File}", Path.GetFileName(sourcePath));

        await using var stream = File.OpenRead(sourcePath);
        using var pkg = new ExcelPackage(stream);

        var ws = pkg.Workbook.Worksheets.FirstOrDefault()
                 ?? throw new InvalidOperationException("Workbook has no worksheets.");

        int rows = ws.Dimension?.Rows ?? 0;
        int cols = ws.Dimension?.Columns ?? 0;

        if (rows < 2)
        {
            _logger.LogWarning("JsonExporter: need at least a header row + data row");
            await File.WriteAllTextAsync(destinationPath, "[]", ct);
            return;
        }

        // Read headers from row 1
        var headers = new string[cols];
        for (int c = 1; c <= cols; c++)
            headers[c - 1] = ws.Cells[1, c].Text?.Trim() is { Length: > 0 } h
                ? h
                : $"Column{c}";

        // Build list of dictionaries
        var records = new List<Dictionary<string, object?>>(rows - 1);

        for (int r = 2; r <= rows; r++)
        {
            var row = new Dictionary<string, object?>(cols);
            for (int c = 1; c <= cols; c++)
            {
                var cell = ws.Cells[r, c];
                row[headers[c - 1]] = InferType(cell);
            }
            records.Add(row);
        }

        var opts = new JsonSerializerOptions
        {
            WriteIndented    = true,
            Encoder          = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        };

        string json = JsonSerializer.Serialize(records, opts);
        await File.WriteAllTextAsync(destinationPath, json, ct);

        _logger.LogInformation(
            "JsonExporter: wrote {Count} records to {Dst}", records.Count, destinationPath);
    }

    private static object? InferType(ExcelRangeBase cell)
    {
        if (cell.Value is null) return null;

        return cell.Value switch
        {
            double d  => d,
            bool   b  => b,
            DateTime dt => dt.ToString("O"),
            _          => cell.Text,
        };
    }
}
