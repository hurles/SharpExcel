using System.Reflection;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel;

/// <summary>
/// This struct is only used to load the metadata of the model
/// </summary>
internal struct PropertyData
{
    public string? Name { get; set; }

    public string? Format { get; set; }

    public int ColumnWidth { get; set; }

    public bool Conditional { get; set; }

    public PropertyInfo PropertyInfo { get; set; }
}

internal class PropertyDataCollection
{
    public Dictionary<Type, List<EnumData>> EnumMappings { get; set; } = new();

    public List<PropertyData> PropertyMappings { get; set; } = new();
}