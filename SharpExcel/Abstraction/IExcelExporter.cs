using ClosedXML.Excel;
using SharpExcel.Models;
using SharpExcel.Models.Results;

namespace SharpExcel.Abstraction;

public interface IExcelExporter<TModel>
    where TModel : class, new()
{
    public Task<XLWorkbook> BuildWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>>? optionalColumnFunc = null);

    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook);

    public Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>>? conditionalColumnFunc = null);

    public HashSet<string> GetOptionalColumns();
}