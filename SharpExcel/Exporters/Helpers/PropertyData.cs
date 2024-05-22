using System.Reflection;

namespace SharpExcel.Exporters.Helpers;

/// <summary>
/// This struct is only used to load the metadata of the model
/// </summary>
internal class PropertyData
{
    
    /// <summary>
    /// Name of property
    /// </summary>
    public string? Name { get; set; }
    
    /// <summary>
    /// Normalized, lowercase name for this property
    /// </summary>
    public string? NormalizedName { get; set; }
    
    /// <summary>
    /// Format string (if applicable)
    /// </summary>
    public string? Format { get; set; }
    
    /// <summary>
    /// Desired width for the column
    /// </summary>
    public int ColumnWidth { get; set; }
    
    /// <summary>
    /// Reflected property info to retrieve property values
    /// </summary>
    public PropertyInfo PropertyInfo { get; set; } = null!;

    /// <summary>
    /// maps address for column
    /// </summary>
    public ColumnData ColumnData { get; set; } = new();
}