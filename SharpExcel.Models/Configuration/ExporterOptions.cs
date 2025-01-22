using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Models.Configuration;

/// <summary>
/// Options to use in the fluent model
/// </summary>
/// <typeparam name="TExportModel">type of model</typeparam>
public class ExporterOptions<TExportModel>
    where TExportModel : class, new()
{
    /// <summary>
    /// Collection of styling rules
    /// </summary>
    public StylingCollection<TExportModel> Styling { get; set; } = new();

    /// <summary>
    /// Fluent method to set default header style for this exporter
    /// </summary>
    /// <param name="style">style to use</param>
    /// <returns></returns>
    public ExporterOptions<TExportModel> WithHeaderStyle(ExcelCellStyle style)
    {
        Styling.DefaultHeaderStyle = style;
        return this;
    }
    
    /// <summary>
    /// Fluent method to set default cell style for this exporter
    /// </summary>
    /// <param name="style">style to use</param>
    /// <returns></returns>
    public ExporterOptions<TExportModel> WithDataStyle(ExcelCellStyle style)
    {
        Styling.DefaultDataStyle = style;
        return this;
    }
    
    /// <summary>
    /// Fluent method to set default error style for this exporter
    /// </summary>
    /// <param name="style">style to use</param>
    /// <returns></returns>
    public ExporterOptions<TExportModel> WithErrorStyle(ExcelCellStyle style)
    {
        Styling.DefaultErrorStyle = style;
        return this;
    }
    
    /// <summary>
    /// Fluent method to add a styling rule for this exporter
    /// </summary>
    /// <param name="stylingRuleOptions">constructs the styling rule</param>
    /// <returns></returns>
    public ExporterOptions<TExportModel> WithStylingRule(Action<StylingRule<TExportModel>> stylingRuleOptions)
    {
        var stylingRule = new StylingRule<TExportModel>();
        stylingRuleOptions(stylingRule);
        Styling.Rules.Add(stylingRule);
        return this;
    }
}