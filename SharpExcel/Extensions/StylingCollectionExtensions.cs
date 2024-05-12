using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Extensions;

internal static class StylingCollectionExtensions
{
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