using System.Reflection;

namespace SharpExcel;

/// <summary>
/// This struct is only used to load the metadata of the model
/// </summary>
internal class PropertyData
{
    public string? Name { get; set; }
    
    public string? NormalizedName { get; set; }

    public string? Format { get; set; }

    public int ColumnWidth { get; set; }
    
    public PropertyInfo PropertyInfo { get; set; } = null!;


    public ColumnData ColumnData { get; set; } = new();
}

internal class PropertyDataCollection
{
    public Dictionary<Type, List<EnumData>> EnumMappings { get; set; } = new();

    public List<PropertyData> PropertyMappings { get; set; } = new();
    
    public Dictionary<string, ColumnData> ColumnMappings { get; set; } = new();
}

internal class ColumnData
{
    public int ColumnIndex { get; set; } = 1;

    public string ColumnName { get; set; } = null!;
}