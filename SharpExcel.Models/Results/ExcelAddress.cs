namespace SharpExcel.Models.Results;

public struct ExcelAddress
{
    public int RowNumber { get; set; }
    public int ColumnId { get; set; }
    public string ColumnName { get; set; }
    public string? HeaderName { get; set; }
}