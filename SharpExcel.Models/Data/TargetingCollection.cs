namespace SharpExcel.Models.Data;

public class TargetingCollection<TExportModel>
{
    public List<TargetingRule<TExportModel>> Rules { get; set; } = new();
}