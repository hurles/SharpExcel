using SharpExcel.Models.Styling.Constants;

namespace SharpExcel.Models.Configuration.Constants;

/// <summary>
/// Constant values for ExporterOptions
/// </summary>
public static class ExporterOptionsConstants
{
    /// <summary>
    /// Default ExporterOptions
    /// </summary>
    /// <typeparam name="TExportModel">type of export model</typeparam>
    /// <returns></returns>
    public static ExporterOptions<TExportModel> GetDefaultOptions<TExportModel>()
        where TExportModel : class, new()
    {
        return new ExporterOptions<TExportModel>()
            .WithDataStyle(SharpExcelCellStyleConstants.DefaultDataStyle)
            .WithHeaderStyle(SharpExcelCellStyleConstants.DefaultHeaderStyle)
            .WithErrorStyle(SharpExcelCellStyleConstants.DefaultErrorStyle);
    }
}