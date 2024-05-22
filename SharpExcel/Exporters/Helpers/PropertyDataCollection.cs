namespace SharpExcel.Exporters.Helpers;

/// <summary>
/// Collection for enum and property mappings
/// </summary>
internal class PropertyDataCollection
{
    /// <summary>
    /// Contains all used enum mappings generated
    /// </summary>
    public Dictionary<Type, List<EnumData>> EnumMappings { get; set; } = new();

    /// <summary>
    /// Contains all property mappings generated from type + it's attributes
    /// </summary>
    public List<PropertyData> PropertyMappings { get; set; } = new();
    
}