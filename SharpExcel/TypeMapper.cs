using System.ComponentModel.DataAnnotations;
using System.Reflection;
using SharpExcel.Models.Attributes;

namespace SharpExcel;

internal static class TypeMapper
{
    /// <summary>
    /// Reads model attributes and converts to column metadata
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public static PropertyDataCollection GetModelMetaData<TModel>() 
    {
        var propertyDataCollection = new PropertyDataCollection();
        
        var dataType = typeof(TModel);
        for (var columnIndex = 0; columnIndex < dataType.GetProperties().Length; columnIndex++)
        {

            var property = dataType.GetProperties()[columnIndex];

            var attribute = property.GetCustomAttribute<ExcelColumnDefinitionAttribute>();

            if (attribute is null)
            {
                continue;
            }

            var columnName = property.Name;
            if (!string.IsNullOrWhiteSpace(attribute?.DisplayName))
            {
                columnName = attribute?.DisplayName;
            }
            
            if (property.PropertyType.IsEnum)
            {
                if (!propertyDataCollection.EnumMappings.ContainsKey(property.PropertyType))
                {
                    propertyDataCollection.EnumMappings.Add(property.PropertyType, GetEnumMappings(property.PropertyType));
                }
            }

            propertyDataCollection.PropertyMappings.Add(new PropertyData()
            {
                Name = columnName,
                NormalizedName = columnName?.ToLowerInvariant(),
                PropertyInfo = property,
                Format = attribute?.Format,
                Conditional = attribute?.IsConditional ?? false,
                ColumnWidth = attribute?.ColumnWidth ?? -1
            });
        }

        return propertyDataCollection;
    }

    private static List<EnumData> GetEnumMappings(Type propertyType)
    {
        if (!propertyType.IsEnum)
            return new();
        
        var enumDataList = new List<EnumData>();
        var memberInfos = propertyType.GetFields();

        foreach (var member in memberInfos.Where(m => m.FieldType.IsEnum))
        {
            var displayAttribute = member.GetCustomAttribute<DisplayAttribute>();
            object enumValue = Enum.ToObject(propertyType, member.GetRawConstantValue());
            
            var enumData = new EnumData()
            {
                VisualName = displayAttribute?.Name ?? member.Name,
                Name = member.Name.ToLowerInvariant(),
                //use long because technically you can use longs for enum values
                NumericValue = Convert.ToInt64(member.GetRawConstantValue()),
                Value = enumValue
            };
            enumDataList.Add(enumData);
        }

        return enumDataList;
    }
}