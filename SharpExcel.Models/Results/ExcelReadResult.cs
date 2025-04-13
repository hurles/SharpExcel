namespace SharpExcel.Models.Results;

public class ExcelReadResult<TModel>
    where TModel : class
{
    public List<TModel> Records  { get; set; } = new();

    public Dictionary<TModel, ExcelCellValidationResult> ValidationResults { get; set; } = new();


}

public static class ExcelReadResultExtensions
{
    public static void Append<TModel>(this ExcelReadResult<TModel> result, ExcelReadResult<TModel> other)
        where TModel : class
    {
        result.Records.AddRange(other.Records);
        foreach (var kvp in other.ValidationResults)
        {
            if (!result.ValidationResults.ContainsKey(kvp.Key))
            {
                result.ValidationResults.Add(kvp.Key, kvp.Value);
                continue;
            }
            result.ValidationResults[kvp.Key] = kvp.Value;
        }

    }
}