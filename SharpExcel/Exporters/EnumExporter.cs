using ClosedXML.Excel;

namespace SharpExcel.Exporters;

internal class EnumExporter
{
    /// <summary>
    /// Fill a hidden worksheet with data from which to fill enum dropdowns with
    /// </summary>
    /// <param name="instance">instance of this run</param>
    /// <returns></returns>
    public static Dictionary<Type, string> AddEnumDropdownMappingsToSheet<TModel>(SharpExcelWriterInstanceData<TModel> instance)
        where TModel : class
    {
        int dropDownWorkbookColumn = 1;
        var dropdownDataMappings = new Dictionary<Type, string>();
        foreach (var enumMapping in instance.Properties.EnumMappings)
        {
            var columnLength = 0;
            for (int i = 0; i < enumMapping.Value.Count; i++)
            {
                var cell = instance.DropdownSourceWorksheet.Row(i + 1).Cell(dropDownWorkbookColumn);
                cell.SetValue(enumMapping.Value[i].VisualName);
                columnLength++;
            }

            var letter = instance.DropdownSourceWorksheet.Column(dropDownWorkbookColumn).ColumnLetter();
            dropdownDataMappings.Add(enumMapping.Key, $"{letter}{1}:{letter}{columnLength}");
            dropDownWorkbookColumn++;
        }

        return dropdownDataMappings;
    }

    /// <summary>
    /// Write enum value into designated cell and add a dropdown with all possible values
    /// </summary>
    /// <param name="instance"></param>
    /// <param name="type"></param>
    /// <param name="dataValue"></param>
    /// <param name="cell"></param>
    /// <param name="dropdownDataMappings"></param>
    public static void WriteEnumValue<TModel>(SharpExcelWriterInstanceData<TModel> instance, Type type, object dataValue,
        IXLCell cell, Dictionary<Type, string> dropdownDataMappings)
    where TModel : class
    {
        if (instance.Properties.EnumMappings.TryGetValue(type, out var enumValues))
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
                        
            if (dropdownDataMappings.TryGetValue(type, out var range))
            {
                cell.CreateDataValidation().List(instance.DropdownSourceWorksheet.Range(range), true);
            }
        }
    }
}