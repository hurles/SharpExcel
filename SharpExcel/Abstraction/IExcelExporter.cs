using ClosedXML.Excel;
using SharpExcel.Models;

namespace SharpExcel.Abstraction;

public interface IExcelExporter<TModel>
    where TModel : class, new()
{
    public Task<XLWorkbook> BuildWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>>? optionalColumnFunc = null);

    public Task<List<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook);

    public Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>>? optionalColumnFunc = null);

    public HashSet<string> GetOptionalColumns();
}