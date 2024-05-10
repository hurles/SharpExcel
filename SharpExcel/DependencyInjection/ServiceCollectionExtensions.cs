using Microsoft.Extensions.DependencyInjection;
using SharpExcel.Abstraction;
using SharpExcel.Exporters;
using SharpExcel.Models.Configuration;
using Microsoft.Extensions.DependencyInjection.Extensions;
using SharpExcel.Models.Configuration.Constants;

namespace SharpExcel.DependencyInjection;

public static class SharpExcelServiceCollectionExtensions
{
    public static void AddDefaultExporter<TExportModel>(this IServiceCollection services)
        where TExportModel : class, new()
    {
        services.Configure<ExporterOptions<TExportModel>>(_ => ExporterOptionsConstants.GetDefaultOptions<TExportModel>());
        services.AddTransient<IExcelExporter<TExportModel>, BaseExcelExporter<TExportModel>>();
    }
    
    public static void AddDefaultExporter<TExportModel>(this IServiceCollection services, Action<ExporterOptions<TExportModel>> options)
        where TExportModel : class, new()
    {
        services.Configure(options);
        services.AddTransient<IExcelExporter<TExportModel>, BaseExcelExporter<TExportModel>>();
    }
    
    public static void AddExporter<TExporter, TExportModel>(this IServiceCollection services, Action<ExporterOptions<TExportModel>> options)
        where TExportModel : class, new()
        where TExporter : BaseExcelExporter<TExportModel>
    {
        services.AddTransient<IExcelExporter<TExportModel>, TExporter>();
    }
}