namespace SharpExcel.Extensions;

/// <summary>
/// Extension methods for types, internally used to determine if types are numeric
/// </summary>
internal static class TypeExtensions
{
    /// <summary>
    /// All possible numeric types
    /// </summary>
    private static readonly HashSet<Type> NumericTypes = new ()
    {
        typeof(int), typeof(double), typeof(decimal),
        typeof(long), typeof(short), typeof(sbyte),
        typeof(byte), typeof(ulong), typeof(ushort),
        typeof(uint), typeof(float)
    };
    
    /// <summary>
    /// Check whether this type is numeric
    /// </summary>
    /// <param name="type">type to check</param>
    /// <returns></returns>
    public static bool IsNumeric(this Type type)
    {
        return NumericTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);
    }
}