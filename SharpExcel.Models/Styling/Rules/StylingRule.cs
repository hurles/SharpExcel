namespace SharpExcel.Models.Styling.Rules;

public class StylingRule<TModel>
{
    public HashSet<string> PropertyNames { get; set; } = new();
    
    private List<Func<TModel, bool>> Conditions { get; set; } = new();
    
    private ExcelCellStyle? StyleWhenTrue { get; set; } = new();
    
    private ExcelCellStyle? StyleWhenFalse { get; set; } = new();
    
    
    public StylingRule<TModel> WithCondition(Func<TModel, bool> condition)
    {
        Conditions.Add(condition);
        //return this object so we can chain calls
        return this;
    }
    
    public StylingRule<TModel> ForProperty(string propertyName)
    {
        PropertyNames.Add(propertyName);
        //return this object so we can chain calls
        return this;
    }
    
    public ExcelCellStyle? EvaluateRules(TModel model)
    {
        bool allAreTrue = false;
        foreach (var condition in Conditions)
        {
            if (!allAreTrue)
            {
                allAreTrue = condition.Invoke(model);
            }
        }

        return allAreTrue ? StyleWhenTrue : StyleWhenFalse;
    }

    public StylingRule<TModel> WhenTrue(ExcelCellStyle style)
    {
        StyleWhenTrue = style;
        //return this object so we can chain calls
        return this;
    }
    
    public StylingRule<TModel> WhenFalse(ExcelCellStyle style)
    {
        StyleWhenFalse = style;
        //return this object so we can chain calls
        return this;
    }
}