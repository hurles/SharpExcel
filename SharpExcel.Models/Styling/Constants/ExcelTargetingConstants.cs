using SharpExcel.Models.Data;

namespace SharpExcel.Models.Styling.Constants;

public class ExcelTargetingConstants<TModel>
    where TModel : class
{
    public static TargetingRule<TModel> DefaultTargetingRule = new TargetingRule<TModel>
    {
        SheetName = "Export",
        Column = 1,
        Row = 1,
        RulePredicate = _ => true,
    };

}