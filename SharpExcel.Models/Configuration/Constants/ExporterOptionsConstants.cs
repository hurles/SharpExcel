using SharpExcel.Models.Styling.Constants;

namespace SharpExcel.Models.Configuration.Constants;

public static class ExporterOptionsConstants
{
    public static ExporterOptions<TExportModel> GetDefaultOptions<TExportModel>()
        where TExportModel : class, new()
    {
        return new ExporterOptions<TExportModel>()
            .WithDataStyle(SharpExcelCellStyleConstants.DefaultDataStyle)
            .WithHeaderStyle(SharpExcelCellStyleConstants.DefaultHeaderStyle)
            .WithErrorStyle(SharpExcelCellStyleConstants.DefaultErrorStyle);
    }
}