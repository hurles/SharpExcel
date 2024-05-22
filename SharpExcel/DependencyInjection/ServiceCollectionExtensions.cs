using Microsoft.Extensions.DependencyInjection;
using SharpExcel.Abstraction;
using SharpExcel.Exporters;
using SharpExcel.Models.Configuration;
using SharpExcel.Models.Configuration.Constants;

namespace SharpExcel.DependencyInjection;

/// <summary>
/// This class contains extension methods for dependency injection
/// </summary>
public static class SharpExcelServiceCollectionExtensions
{
    /// <summary>
    /// Adds a default exporter type to the ServiceCollection, with the provided model type
    /// </summary>
    /// <param name="services">the ServiceCollection</param>
    /// <param name="options">the options to configure this SharpExcel exporter with. Default options are used when null</param>
    /// <typeparam name="TExportModel"></typeparam>
    public static void AddSynchronizer<TExportModel>(this IServiceCollection services, Action<ExporterOptions<TExportModel>>? options)
        where TExportModel : class, new()
    {
        services.Configure(options ?? (_ => ExporterOptionsConstants.GetDefaultOptions<TExportModel>()));
        services.AddTransient<ISharpExcelSynchronizer<TExportModel>, BaseSharpExcelSynchronizer<TExportModel>>();
    }
    
    /// <summary>
    /// Adds a customized exporter type to the ServiceCollection, with the provided model type.
    /// </summary>
    /// <param name="services">the ServiceCollection</param>
    /// <param name="options">the options to configure this SharpExcel exporter with. Default options are used when null</param>
    /// <typeparam name="TExporter">The type of the exporter. Must be inherited from BaseExcelExporter </typeparam>
    /// <typeparam name="TExportModel">any class to use as data model</typeparam>
    public static void AddSynchronizer<TExporter, TExportModel>(this IServiceCollection services, Action<ExporterOptions<TExportModel>>? options)
        where TExportModel : class, new()
        where TExporter : BaseSharpExcelSynchronizer<TExportModel>
    {
        
        services.Configure(options ?? (_ => ExporterOptionsConstants.GetDefaultOptions<TExportModel>()));
        services.AddTransient<ISharpExcelSynchronizer<TExportModel>, TExporter>();
    }
}