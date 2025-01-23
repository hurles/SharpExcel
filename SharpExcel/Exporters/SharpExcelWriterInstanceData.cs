using System.Globalization;
using ClosedXML.Excel;
using SharpExcel.Exporters.Helpers;
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
    /// <summary>
    /// Header style to use for this instance
    /// </summary>
    public ExcelCellStyle HeaderStyle { get; set; }
    
    /// <summary>
    /// Error style to use for this instance
    /// </summary>
    public ExcelCellStyle ErrorStyle { get; set; }
    
    /// <summary>
    /// Style to use for data cells
    /// </summary>
    public ExcelCellStyle DataStyle { get; set; }
    
    /// <summary>
    /// lookup for styling rules, so we can look up rules for properties faster
    /// </summary>
    public Dictionary<string, List<StylingRule<TModel>>> StylingRuleLookup { get; set; } = new();
    
    /// <summary>
    /// Collection of property metadata (column name etc.)
    /// </summary>
    public PropertyDataCollection Properties { get; set; } = new();
    
    /// <summary>
    /// Main worksheet to use for reading/writing
    /// </summary>
    public IXLWorksheet MainWorksheet { get; set; } = null!;
    
    /// <summary>
    /// Hidden worksheet to serve as source for all generated dropdown menus
    /// </summary>
    public IXLWorksheet DropdownSourceWorksheet { get; set; } = null!;
    
    /// <summary>
    /// Culture info used for reading/writing in this instance
    /// </summary>
    public CultureInfo? CultureInfo { get; set; } = null!;
}