using SharpExcel.Models.Styling.Constants;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Models.Styling;

public class StylingCollection<TExportModel>
{
    public ExcelCellStyle DefaultHeaderStyle { get; set; } = ExcelCellStyleConstants.DefaultHeaderStyle;

    public ExcelCellStyle DefaultDataStyle { get; set; } = ExcelCellStyleConstants.DefaultDataStyle;
    
    
    public ExcelCellStyle DefaultErrorStyle { get; set; } = ExcelCellStyleConstants.DefaultErrorStyle;

    public List<StylingRule<TExportModel>> Rules { get; set; } = new();
}