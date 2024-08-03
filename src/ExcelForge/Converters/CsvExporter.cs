using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Text;

namespace ExcelForge.Converters;

/// <summary>
/// Exports the first worksheet of an Excel file to CSV (UTF-8 with BOM).
/// </summary>
internal sealed class CsvExporter : IConverter
{
    private readonly ILogger<CsvExporter> _logger;

    public CsvExporter(ILogger<CsvExporter> logger)
    {
        _logger = logger;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task ConvertAsync(string sourcePath, string destinationPath,
                                   CancellationToken ct = default)
    {
        _logger.LogInformation("CsvExporter: reading {File}", Path.GetFileName(sourcePath));

        await using var stream = File.OpenRead(sourcePath);
        using var pkg = new ExcelPackage(stream);

        var ws = pkg.Workbook.Worksheets.FirstOrDefault()
                 ?? throw new InvalidOperationException("Workbook has no worksheets.");

        int rows = ws.Dimension?.Rows ?? 0;
        int cols = ws.Dimension?.Columns ?? 0;

        if (rows == 0)
        {
            _logger.LogWarning("CsvExporter: worksheet '{Name}' is empty", ws.Name);
            await File.WriteAllTextAsync(destinationPath, string.Empty, ct);
            return;
        }

        var sb = new StringBuilder(capacity: rows * cols * 12);

        for (int r = 1; r <= rows; r++)
        {
            for (int c = 1; c <= cols; c++)
            {
                var val = ws.Cells[r, c].Text ?? string.Empty;

                // RFC 4180 quoting
                if (val.Contains(',') || val.Contains('"') || val.Contains('\n'))
                {
                    val = "\"" + val.Replace("\"", "\"\"") + "\"";
                }

                sb.Append(val);
                if (c < cols) sb.Append(',');
            }
            sb.Append('\n');
        }

        await File.WriteAllTextAsync(destinationPath, sb.ToString(),
                                     new UTF8Encoding(encoderShouldEmitUTF8Identifier: true), ct);

        _logger.LogInformation("CsvExporter: wrote {Rows} rows to {Dst}", rows, destinationPath);
    }
}
