namespace SharpExcel.Extensions;

internal static class TypeExtensions
{
    private static readonly HashSet<Type> NumericTypes = new ()
    {
        typeof(int), typeof(double), typeof(decimal),
        typeof(long), typeof(short), typeof(sbyte),
        typeof(byte), typeof(ulong), typeof(ushort),
        typeof(uint), typeof(float)
    };
    
    public static bool IsNumeric(this Type type)
    {
        return NumericTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);
    }
}