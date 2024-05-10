using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Models.Arguments.Extensions;

public static class ExcelArgumentsExtensions
{
    public static StylingRule<TModel> AddStylingRule<TModel>(this ExcelArguments<TModel> arguments)
        where TModel : class
    {
        var stylingRule = new StylingRule<TModel>();
        arguments.StylingCollection.Rules.Add(stylingRule);
        return stylingRule;
    }
}