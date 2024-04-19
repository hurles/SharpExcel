namespace SharpExcel.Attributes;

public class ExcelColumnDefinitionAttribute(
    string columnName,
    int width = -1,
    string? format = null,
    bool isOptional = false)
    : Attribute
{
    public string DisplayName { get; set; } = columnName;

    public int ColumnWidth { get; set; } = width;

    public string? Format { get; set; } = format;

    public bool IsOptional { get; set; } = isOptional;
}