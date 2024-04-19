using ClosedXML.Excel;
using ExcelSharp.Models;

namespace ExcelSharp.Abstraction;

public interface IExcelExporter<TModel>
    where TModel : IExcelModel, new()
{
    public Task<XLWorkbook> BuildWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>> optionalColumnFunc = null);

    public Task<List<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook);

    public Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>> optionalColumnFunc = null);

    public HashSet<string> GetOptionalColumns();
}