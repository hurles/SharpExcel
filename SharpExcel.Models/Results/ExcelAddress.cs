namespace SharpExcel.Models.Results;

/// <summary>
/// Data structure for addressing a single cell
/// </summary>
public record struct ExcelAddress
{
    /// <summary>
    /// The row number
    /// </summary>
    public int RowNumber { get; set; }
    
    /// <summary>
    /// Column id
    /// </summary>
    public int ColumnId { get; set; }
    
    /// <summary>
    /// Column Name
    /// </summary>
    public string ColumnName { get; set; }
    
    /// <summary>
    /// Header name
    /// </summary>
    public string? HeaderName { get; set; }
    
    /// <summary>
    /// Sheet Name
    /// </summary>
    public string SheetName { get; set; }
}