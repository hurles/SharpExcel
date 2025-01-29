using SharpExcel.Models.Targeting;

namespace SharpExcel.Extensions;

public static class TargetRuleCollectionExtensions
{
    /// <summary>
    /// Adds a rule to the collection
    /// </summary>
    /// <param name="collection">The collection to modify</param>
    /// <param name="rule">The rule to be added</param>
    /// <typeparam name="TRule">Type parameter for rule</typeparam>
    /// <returns>The collection, to chain other calls</returns>
    public static ExcelTargetRuleCollection WithRule<TRule>(this ExcelTargetRuleCollection collection, TRule rule)
        where TRule : SharpExcelCellTargetRule
    {
        collection.TargetRules.Add(rule);
        return collection;
    }
}