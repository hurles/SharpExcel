using Microsoft.Extensions.DependencyInjection;
using SharpExcel.Abstraction;
using SharpExcel.Exporters;

namespace SharpExcel.DependencyInjection;

public static class SharpExcelServiceCollectionExtensions
{
    public static void AddDefaultExporter<TModel>(this IServiceCollection services)
        where TModel : class, new()
    {
        services.AddTransient<IExcelExporter<TModel>, BaseExcelExporter<TModel>>();
    }
    
    public static void AddExporter<TExporter, TModel>(this IServiceCollection services)
        where TModel : class, new()
        where TExporter : BaseExcelExporter<TModel>
    {
        services.AddTransient<IExcelExporter<TModel>, TExporter>();
    }
}