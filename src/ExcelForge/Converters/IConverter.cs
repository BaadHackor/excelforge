namespace ExcelForge.Converters;

/// <summary>
/// Common interface for all file converters.
/// </summary>
internal interface IConverter
{
    /// <summary>
    /// Converts <paramref name="sourcePath"/> and writes the result to
    /// <paramref name="destinationPath"/>. Throws on error.
    /// </summary>
    Task ConvertAsync(string sourcePath, string destinationPath,
                      CancellationToken ct = default);
}
