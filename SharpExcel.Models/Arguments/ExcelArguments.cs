using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Models.Arguments;

public class ExcelArguments<TExportModel>
    where TExportModel : class
{
    public string? SheetName { get; set; }

    public List<TExportModel> Data { get; set; } = new();
    
    
    public StylingCollection<TExportModel> StylingCollection = new();

}

public class StylingCollection<TExportModel>
{
    public SharpExcelCellStyle DefaultHeaderStyle { get; set; } = SharpExcelCellStyleConstants.DefaultHeaderStyle;

    public SharpExcelCellStyle DefaultDataStyle { get; set; } = SharpExcelCellStyleConstants.DefaultDataStyle;
    
    public SharpExcelCellStyle DefaultErrorStyle { get; set; } = SharpExcelCellStyleConstants.DefaultErrorStyle;

    public List<StylingRule<TExportModel>> Rules { get; set; } = new();
}