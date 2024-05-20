using System.Globalization;
using ClosedXML.Excel;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Exporters;

/// <summary>
/// Holds several lookups so we can reuse data during construction of workbooks
/// </summary>
/// <typeparam name="TModel">type of model</typeparam>
internal class SharpExcelWriterInstanceData<TModel> 
    where TModel : class
{
    public SharpExcelCellStyle HeaderStyle { get; set; }
    public SharpExcelCellStyle ErrorStyle { get; set; }
    public SharpExcelCellStyle DataStyle { get; set; }
    public Dictionary<string, List<StylingRule<TModel>>> StylingRuleLookup { get; set; } = new();
    public PropertyDataCollection Properties { get; set; } = new();
    public IXLWorksheet MainWorksheet { get; set; } = null!;
    public IXLWorksheet DropdownSourceWorksheet { get; set; } = null!;
    public CultureInfo? CultureInfo { get; set; } = null!;
}