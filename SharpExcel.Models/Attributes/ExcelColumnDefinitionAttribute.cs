namespace SharpExcel.Models.Attributes;

/// <summary>
/// This attribute maps the property to a column in Excel
/// </summary>
/// <param name="columnName">Name for the column in Excel</param>
/// <param name="width">column width to use (-1 means no value (default))</param>
/// <param name="format">optional format string (this will convert the value into a string in Excel)</param>
public class ExcelColumnDefinitionAttribute(string columnName, int width = -1, string? format = null)
    : Attribute
{
    public string DisplayName { get; set; } = columnName;

    public int ColumnWidth { get; set; } = width;

    public string? Format { get; set; } = format;
}