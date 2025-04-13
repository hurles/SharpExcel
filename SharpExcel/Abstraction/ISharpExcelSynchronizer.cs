using System.Globalization;
using ClosedXML.Excel;
using SharpExcel.Models.Results;

namespace SharpExcel.Abstraction;

/// <summary>
/// Main interface for excel exports and imports
/// </summary>
/// <typeparam name="TModel"></typeparam>
public interface ISharpExcelSynchronizer<TModel>
    where TModel : class, new()
{
    /// <summary>
    /// Generates a workbook based on the provided data
    /// </summary>
    /// <param name="cultureInfo"></param>
    /// <param name="data">The data to generate the workbook from</param>
    /// <returns></returns>
    public Task<XLWorkbook> GenerateWorkbookAsync(CultureInfo cultureInfo, ICollection<TModel> data);

    /// <summary>
    /// Reads a workbook to convert it into the given model
    /// </summary>
    /// <param name="arguments">Collection of arguments</param>
    /// <param name="workbook"></param>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(CultureInfo arguments, XLWorkbook workbook);

    /// <summary>
    /// Reads, then returns the supplied workbook, but highlights cells containing invalid data, using standard System.ComponentModel.DataAnnotations validation on the model
    /// </summary>
    /// <param name="cultureInfo"></param>
    /// <param name="workbook">The workbook</param>
    /// <returns>The highlighted workbook</returns>
    public Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(CultureInfo cultureInfo, XLWorkbook workbook);
}