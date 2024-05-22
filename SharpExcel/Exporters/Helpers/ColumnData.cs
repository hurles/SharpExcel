namespace SharpExcel.Exporters.Helpers;

/// <summary>
/// Mapping data for columnss and their indexes, so we can read the right column for the right property
/// </summary>
internal class ColumnData
{
    /// <summary>
    /// Column index
    /// </summary>
    public int ColumnIndex { get; set; } = 1;

    /// <summary>
    /// Column name derived from property attribute/name
    /// </summary>
    public string ColumnName { get; set; } = null!;
}