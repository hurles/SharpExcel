
namespace SharpExcel.Exporters.Helpers;

/// <summary>
/// Mapping data for enums + their display names
/// </summary>
internal struct EnumData
{
    /// <summary>
    /// Display name to use in excel files for this enum member
    /// </summary>
    public string? VisualName { get; set; }
    
    /// <summary>
    /// Name of enum member
    /// </summary>
    public string? Name { get; set; }
    
    /// <summary>
    /// Numeric value of this enum member
    /// </summary>
    public long NumericValue { get; set; }

    /// <summary>
    /// value of this enum member
    /// </summary>
    public object Value { get; set; }
}