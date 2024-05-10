
namespace SharpExcel;

internal struct EnumData
{
    public string? VisualName { get; set; }
    public string? Name { get; set; }
    
    //use long because technically you can use longs for enum values
    public long NumericValue { get; set; }

    public object Value { get; set; }
}