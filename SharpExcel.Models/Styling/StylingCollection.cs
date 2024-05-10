using SharpExcel.Models.Styling.Constants;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Models.Styling;

public class StylingCollection<TExportModel>
{
    public SharpExcelCellStyle DefaultHeaderStyle { get; set; } = SharpExcelCellStyleConstants.DefaultHeaderStyle;

    public SharpExcelCellStyle DefaultDataStyle { get; set; } = SharpExcelCellStyleConstants.DefaultDataStyle;
    
    
    public SharpExcelCellStyle DefaultErrorStyle { get; set; } = SharpExcelCellStyleConstants.DefaultErrorStyle;

    public List<StylingRule<TExportModel>> Rules { get; set; } = new();
}