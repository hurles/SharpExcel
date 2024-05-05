using System.Globalization;
using ClosedXML.Excel;
using SharpExcel.Models;
using SharpExcel.Models.Results;

namespace SharpExcel.Abstraction;

public interface IExcelExporter<TModel>
    where TModel : class, new()
{
    public Task<XLWorkbook> GenerateWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>>? conditionalColumnFunc = null, CultureInfo? cultureInfo = null);

    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null);

    public Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>>? conditionalColumnFunc = null);
}