namespace SharpExcel.Models.Results;

/// <summary>
/// Data structure 
/// </summary>
public struct ExcelAddress
{
    public int RowNumber { get; set; }
    public int ColumnId { get; set; }
    public string ColumnName { get; set; }
    public string? HeaderName { get; set; }
}