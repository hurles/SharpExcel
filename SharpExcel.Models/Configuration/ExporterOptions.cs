using Microsoft.Extensions.Options;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Models.Configuration;

public class ExporterOptions<TExportModel>
    where TExportModel : class, new()
{
    public StylingCollection<TExportModel> Styling = new();

    public ExporterOptions<TExportModel> WithHeaderStyle(SharpExcelCellStyle style)
    {
        Styling.DefaultHeaderStyle = style;
        return this;
    }
    
    public ExporterOptions<TExportModel> WithDataStyle(SharpExcelCellStyle style)
    {
        Styling.DefaultDataStyle = style;
        return this;
    }
    
    public ExporterOptions<TExportModel> WithErrorStyle(SharpExcelCellStyle style)
    {
        Styling.DefaultErrorStyle = style;
        return this;
    }

    public StylingRule<TExportModel> AddStylingRule()
    {
        var stylingRule = new StylingRule<TExportModel>();
        Styling.Rules.Add(stylingRule);
        return stylingRule;
    }
        
        
}