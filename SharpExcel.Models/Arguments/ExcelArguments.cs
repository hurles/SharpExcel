using System.Globalization;

namespace SharpExcel.Models.Arguments;

/// <summary>
/// Arguments used for reading/writing excel files
/// </summary>
public class ExcelArguments
{
    /// <summary>
    /// Name of Excel worksheet to read/write
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// CultureInfo to use when reading/writing
    /// </summary>
    public CultureInfo? CultureInfo { get; set; }
}