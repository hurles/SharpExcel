using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Extensions;

/// <summary>
/// Extension methods for StylingCollection
/// </summary>
internal static class StylingCollectionExtensions
{
    /// <summary>
    /// Creates lookup for styling rules, used internally to speed up looking up styling rules
    /// </summary>
    /// <param name="collection">the collection to create a lookup from</param>
    /// <typeparam name="TExportModel">Model type</typeparam>
    /// <returns></returns>
    public static Dictionary<string, List<StylingRule<TExportModel>>> ToStylingRuleLookup<TExportModel>(this StylingCollection<TExportModel> collection)
        where TExportModel : class
    {
        var lookup = new Dictionary<string, List<StylingRule<TExportModel>>>();
        foreach (var rule in collection.Rules)
        {
            foreach (var property in rule.PropertyNames)
            {
                if (!lookup.ContainsKey(property))
                {
                    lookup[property] = new List<StylingRule<TExportModel>>();
                }

                //add rule multiple times if it's used for more than one property
                lookup[property].Add(rule);
            }
        }

        return lookup;
    }
}