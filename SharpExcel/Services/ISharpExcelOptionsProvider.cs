using SharpExcel.Models.Configuration;

namespace SharpExcel.Services;

public interface ISharpExcelOptionsProvider<TExportModel>
    where TExportModel : class, new()
{
    public ExporterOptions<TExportModel> GetSharpExcelConfiguration();
}