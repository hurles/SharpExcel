using System.Globalization;

namespace SharpExcel.Models.Arguments;

public class SharpExcelArguments
{
    public string? SheetName { get; set; }

    public CultureInfo? CultureInfo { get; set; }
}