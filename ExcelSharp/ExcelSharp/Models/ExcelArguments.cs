using ExcelSharp.Abstraction;

namespace ExcelSharp.Models;

public class ExcelArguments<TExportModel>
    where TExportModel : class
{
    public string? SheetName { get; set; }

    public List<TExportModel> Data { get; set; } = new();
}