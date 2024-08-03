using ExcelForge.Native;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace ExcelForge.Converters;

/// <summary>
/// Converts .xls / .xlsx workbooks to .xlsx format.
/// Uses EPPlus for pure .xlsx files; delegates legacy .xls handling to the
/// native COM bridge (ExcelForge.Native.dll) when available.
/// </summary>
internal sealed class XlsxConverter : IConverter
{
    private readonly ILogger<XlsxConverter> _logger;

    public XlsxConverter(ILogger<XlsxConverter> logger)
    {
        _logger = logger;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task ConvertAsync(string sourcePath, string destinationPath,
                                   CancellationToken ct = default)
    {
        string ext = Path.GetExtension(sourcePath).ToLowerInvariant();

        if (ext == ".xls")
        {
            await ConvertLegacyAsync(sourcePath, destinationPath, ct);
        }
        else
        {
            await ConvertModernAsync(sourcePath, destinationPath, ct);
        }
    }

    // ── .xlsx → .xlsx (normalize / clean) ────────────────────────────────────

    private async Task ConvertModernAsync(string src, string dst, CancellationToken ct)
    {
        _logger.LogInformation("XlsxConverter: normalizing {File}", Path.GetFileName(src));

        await using var stream = File.OpenRead(src);
        using var pkg = new ExcelPackage(stream);

        // Strip volatile formula caches so the output is deterministic
        foreach (var ws in pkg.Workbook.Worksheets)
        {
            foreach (var cell in ws.Cells)
                cell.Calculate();
        }

        await pkg.SaveAsAsync(new FileInfo(dst), ct);

        _logger.LogInformation("XlsxConverter: wrote {Dst}", Path.GetFileName(dst));
    }

    // ── .xls → .xlsx (via native bridge) ─────────────────────────────────────

    private async Task ConvertLegacyAsync(string src, string dst, CancellationToken ct)
    {
        _logger.LogInformation(
            "XlsxConverter: legacy .xls detected — using native bridge for {File}",
            Path.GetFileName(src));

        using var bridge = new NativeBridge(_logger);

        if (!bridge.IsAvailable)
        {
            _logger.LogWarning(
                "XlsxConverter: native bridge unavailable — falling back to EPPlus read-only mode");
            await FallbackReadAsync(src, dst, ct);
            return;
        }

        // NativeBridge.Convert is blocking (COM STA), so run on thread pool
        await Task.Run(() =>
        {
            bridge.Init();
            int result = bridge.Convert(src, dst, flags: 0x01);
            if (result != 0)
                throw new InvalidOperationException(
                    $"ExcelNative_Convert failed with code 0x{result:X8}");
        }, ct);

        _logger.LogInformation("XlsxConverter: native conversion complete → {Dst}", dst);
    }

    private async Task FallbackReadAsync(string src, string dst, CancellationToken ct)
    {
        // EPPlus can read .xls with limited fidelity
        await using var stream = File.OpenRead(src);
        using var pkg = new ExcelPackage(stream);
        await pkg.SaveAsAsync(new FileInfo(dst), ct);
    }
}
