namespace SharpExcel.Models.Results;

public class ExcelReadResult<TModel>
    where TModel : class
{
    public List<TModel> Records  { get; set; } = new();

    public Dictionary<TModel, ExcelCellValidationResult> ValidationResults { get; set; } = new();
}