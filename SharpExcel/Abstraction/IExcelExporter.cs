using System.Globalization;
using ClosedXML.Excel;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Results;

namespace SharpExcel.Abstraction;

public interface IExcelExporter<TModel>
    where TModel : class, new()
{

    /// <summary>
    /// Generates a workbook based on the provided data
    /// </summary>
    /// <param name="arguments">Collection of arguments</param>
    /// <param name="optionalColumnFunc">Functions that returns a boolean based on Property name of TModel (not column name), indicating whether or not to write the specified property based on a condition</param>
    /// <param name="cultureInfo">Culture used to generate workbook</param>
    /// <returns></returns>
    public Task<XLWorkbook> GenerateWorkbookAsync(SharpExcelArguments arguments, IEnumerable<TModel> data, CultureInfo? cultureInfo = null);

    /// <summary>
    /// Reads a workbook to convert it into the given model
    /// </summary>
    /// <param name="sheetName">name of the sheet to read from</param>
    /// <param name="workbook"></param>
    /// <param name="cultureInfo">culture used, defaults to CurrentCulture if null</param>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null);

    /// <summary>
    /// Reads, then returns the supplied workbook, but highlights cells containing invalid data, using standard System.ComponentModel.DataAnnotations validation on the model
    /// </summary>
    /// <param name="sheetName">Name of the sheet to analyze</param>
    /// <param name="workbook">The workbook</param>
    /// <param name="cultureInfo">Optional culture used for reading and writing</param>
    /// <returns>The highlighted workbook</returns>
    public Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null);
}