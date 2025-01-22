using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using SharpExcel.Exporters.Helpers;
using SharpExcel.Extensions;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Constants;

namespace SharpExcel.Exporters;

/// <summary>
/// Helper methods for exporters
/// </summary>
internal static class ExporterHelpers
{
    /// <summary>
    /// Applies style and adds validation comments to invalid cells
    /// </summary>
    /// <param name="sheetName"></param>
    /// <param name="workbook"></param>
    /// <param name="parsedWorkbook"></param>
    /// <typeparam name="TModel"></typeparam>
    public static void ApplyCellValidation<TModel>(string sheetName, XLWorkbook workbook, ExcelReadResult<TModel> parsedWorkbook)
        where TModel : class, new()
    {
        foreach (var result in parsedWorkbook.ValidationResults)
        {
            var cell = workbook.Worksheet(sheetName).Cell(result.Value.Address.RowNumber, result.Value.Address.ColumnId);
            var stringBuilder = new StringBuilder();
            foreach (var item in result.Value.ValidationResults)
            {
                stringBuilder.AppendLine(item.ErrorMessage);
            }
            cell.Style.ApplyStyle(ExcelCellStyleConstants.DefaultErrorStyle);
            cell.CreateComment().AddText(stringBuilder.ToString());
        }
    }
    
    /// <summary>
    /// Tries to set value according to the current property's type
    /// </summary>
    /// <param name="enumMappings">enum mappings</param>
    /// <param name="columnData">property data for this cell</param>
    /// <param name="cell">cell to wrtie to</param>
    /// <param name="cultureInfo"></param>
    /// <returns></returns>
    public static object? TrySetCellValue(Dictionary<Type, List<EnumData>> enumMappings, PropertyData columnData, IXLCell cell, CultureInfo cultureInfo)
    {
        //extract underlying nullable type if there is one
        var actualType = Nullable.GetUnderlyingType(columnData.PropertyInfo.PropertyType) ?? columnData.PropertyInfo.PropertyType;
        
        //handle numeric types
        if (actualType.IsNumeric())
        {
            if (cell.TryGetValue(out decimal numericValue))
            {
                return Convert.ChangeType(numericValue, actualType);
            }
        }

        //handle booleans
        if (actualType == typeof(bool))
        {
            if (cell.TryGetValue(out bool booleanValue))
            {
                return booleanValue;
            }
        }

        //handle strings
        if (actualType == typeof(string))
        {
            if (cell.TryGetValue(out string textValue))
            {
                return textValue;
            }

            var displayVal = cell.CachedValue.ToString(cultureInfo);
            
            if (!string.IsNullOrWhiteSpace(displayVal))
            {
                return displayVal;
            }
        }

        if (actualType.IsEnum)
        {
            if (cell.TryGetValue(out string textValue))
            {
                if (enumMappings.TryGetValue(actualType, out var data))
                {
                    var value = data.FirstOrDefault(x => x.VisualName?.ToLowerInvariant() == textValue.Trim().ToLowerInvariant());
                    return value.Value;
                }
            }
            
        }
        
        return default;
    }
    
    /// <summary>
    /// Writes a row of data cells
    /// </summary>
    /// <param name="instance">instance data for this run</param>
    /// <param name="dataItem">the current data item being processed</param>
    /// <param name="rowIndex">index of the row to write to</param>
    /// <param name="dropdownDataMappings">dropdown data mappings for the hidden enum dropdown sheet</param>
    /// <typeparam name="TModel">type of the data item being processed</typeparam>
    public static void WriteDataRow<TModel>(
        SharpExcelWriterInstanceData<TModel> instance, 
        TModel dataItem,
        int rowIndex, 
        Dictionary<Type, string> dropdownDataMappings)
        where TModel : class
    {
        
        for (var i = 0; i < instance.Properties.PropertyMappings.Count; i++)
        {
            var mapping = instance.Properties.PropertyMappings[i];
            var row = instance.MainWorksheet.Row(rowIndex);
            var cell = instance.MainWorksheet.Cell(rowIndex, i + 1 /* use +1 because Excel starts at 1 */);
            WriteDataCell(instance, dataItem, dropdownDataMappings, mapping, cell);
            cell.Style.ApplyStyle(GetCellStyle(instance, dataItem, mapping, row));
        }
    }

    /// <summary>
    /// Writes data to a cell
    /// </summary>
    /// <param name="instance">instance data for this run</param>
    /// <param name="dataItem">the current data item being processed</param>
    /// <param name="dropdownDataMappings">dropdown data mappings for the hidden enum dropdown sheet</param>
    /// <param name="mapping"></param>
    /// <param name="cell"></param>
    /// <typeparam name="TModel">type of the data item being processed</typeparam>
    private static void WriteDataCell<TModel>(
        SharpExcelWriterInstanceData<TModel> instance,
        TModel dataItem, 
        Dictionary<Type, string> dropdownDataMappings,
        PropertyData mapping, 
        IXLCell cell)
        where TModel : class
    {
        var dataValue = mapping.PropertyInfo.GetValue(dataItem);
                
        //handle enums
        if (mapping.PropertyInfo.PropertyType.IsEnum)
        {
            EnumExporter.WriteEnumValue(instance, mapping.PropertyInfo.PropertyType, dataValue, cell, dropdownDataMappings);
        }
        //handle format
        else if (mapping.Format != null)
        {
            if (dataValue is IFormattable formattable)
            {
                cell.SetValue(formattable.ToString(mapping.Format, CultureInfo.InvariantCulture));
            }
        }
        else
        {
            cell.SetValue(XLCellValue.FromObject(dataValue));
        }
    }

    /// <summary>
    /// Gets cell style for currently parsed cell, applying styling rules, if any
    /// </summary>
    /// <param name="instance">instance data for this run</param>
    /// <param name="dataItem">the current data item being processed</param>
    /// <param name="mapping">mapping data for the current property</param>
    /// <param name="row">row to apply row height to if necessary</param>
    /// <typeparam name="TModel">type of the data item being processed</typeparam>
    /// <returns></returns>
    private static ExcelCellStyle GetCellStyle<TModel>(SharpExcelWriterInstanceData<TModel> instance, TModel dataItem, PropertyData mapping, IXLRow row) 
        where TModel : class
    {
        var dataStyle = instance.DataStyle;
        
        if (instance.StylingRuleLookup.TryGetValue(mapping.PropertyInfo.Name, out var rules))
        {
            foreach (var rule in rules)
            {
                var ruleStyle = rule.EvaluateRules(dataItem);
                dataStyle = ruleStyle ?? dataStyle;
            }
        }

        if (dataStyle.RowHeight.HasValue && row.Height < dataStyle.RowHeight)
        {
            row.Height = dataStyle.RowHeight.Value;
        }
        
        return dataStyle;
    }

    /// <summary>
    /// Writes header row to cell
    /// </summary>
    /// <param name="instance">instance data for this run</param>
    /// <param name="rowIndex">the row index to write to</param>
    /// <typeparam name="TModel">type of the data item being processed</typeparam>
    public static void WriteHeaderRow<TModel>(
        SharpExcelWriterInstanceData<TModel> instance,  
        int rowIndex)
        where TModel : class
    {
        for (var columnIndex = 0; columnIndex < instance.Properties.PropertyMappings.Count; columnIndex++)
        {
            var mapping = instance.Properties.PropertyMappings[columnIndex];
            
            var cell = instance.MainWorksheet.Cell(rowIndex, columnIndex + 1 /* use +1 because Excel starts at 1 */);
            cell.Style.ApplyStyle(instance.HeaderStyle);

            if (mapping.ColumnWidth > 0)
            {
                instance.MainWorksheet.Column(columnIndex + 1).Width = mapping.ColumnWidth;
            }

            cell.SetValue(mapping.Name);
        }
    }
}