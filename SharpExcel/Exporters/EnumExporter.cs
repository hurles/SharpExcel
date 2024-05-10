using ClosedXML.Excel;

namespace SharpExcel.Exporters;

internal class EnumExporter
{
    /// <summary>
    /// Fill a hidden worksheet with data from which to fill enum dropdowns with
    /// </summary>
    /// <param name="propertyMappings">Pre-parsed property data for the current model</param>
    /// <param name="dropdownWorksheet">Reference to the worksheet to add the dropdown data to</param>
    /// <returns></returns>
    public static Dictionary<Type, string> AddEnumDropdownMappings(PropertyDataCollection propertyMappings, IXLWorksheet dropdownWorksheet)
    {
        int dropDownWorkbookColumn = 1;
        var dropdownDataMappings = new Dictionary<Type, string>();
        foreach (var enumMapping in propertyMappings.EnumMappings)
        {
            var columnLength = 0;
            for (int i = 0; i < enumMapping.Value.Count; i++)
            {
                var cell = dropdownWorksheet.Row(i + 1).Cell(dropDownWorkbookColumn);
                cell.SetValue(enumMapping.Value[i].VisualName);
                columnLength++;
            }

            var letter = dropdownWorksheet.Column(dropDownWorkbookColumn).ColumnLetter();
            dropdownDataMappings.Add(enumMapping.Key, $"{letter}{1}:{letter}{columnLength}");
            dropDownWorkbookColumn++;
        }

        return dropdownDataMappings;
    }
    
    /// <summary>
    /// Write enum value into designated cell and add a dropdown with all possible values
    /// </summary>
    /// <param name="propertyMappings"></param>
    /// <param name="mapping"></param>
    /// <param name="dataValue"></param>
    /// <param name="cell"></param>
    /// <param name="dropdownDataMappings"></param>
    /// <param name="dropdownWorksheet"></param>
    public static void WriteEnumValue(PropertyDataCollection propertyMappings, PropertyData mapping, object dataValue,
        IXLCell cell, Dictionary<Type, string> dropdownDataMappings, IXLWorksheet dropdownWorksheet)
    {
        if (propertyMappings.EnumMappings.TryGetValue(mapping.PropertyInfo.PropertyType, out var enumValues))
        {
            var text = dataValue.ToString().Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(text))
            {
                foreach (var enumValue in enumValues)
                {
                    if (enumValue.Name == text)
                    {
                        cell.SetValue(enumValue.VisualName);
                    }
                }
            }
                        
            if (dropdownDataMappings.TryGetValue(mapping.PropertyInfo.PropertyType, out var range))
            {
                cell.CreateDataValidation().List(dropdownWorksheet.Range(range), true);
            }
        }
    }
}